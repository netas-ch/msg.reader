/* Copyright 2021 Yury Karpovich
 * Modified 2023 by Lukas Buchs, netas.ch. Changed to javascript export class.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
import {DataStream} from './DataStream.js';

/*
 MSG Reader
 */
export class MsgReader {
    #ds;
    #fileData;
    #headers;

    constructor(arrayBuffer) {
        this.#ds = new DataStream(arrayBuffer, 0, DataStream.LITTLE_ENDIAN);

        if (!MsgReader.#isMSGFile(this.#ds)) {
            throw new Error('Unsupported file type!');
        }

        this.#fileData = MsgReader.#parseMsgData(this.#ds);

        if (this.#fileData.fieldsData && this.#fileData.fieldsData.TransportMessageHeaders) {
            this.#headers = MsgReader.#splitHeaders(this.#fileData.fieldsData.TransportMessageHeaders);
        }
    }

    // ----------------------------
    // PUBLIC FUNCTIONS
    // ----------------------------

    /**
     Converts bytes to fields information
     @return {Object} The fields data for MSG file
     */
    getFileData() {
        if (this.#fileData) {
            return this.#fileData.fieldsData;
        }
        return null;
    }

    /**
     Reads an attachment content by key/ID
     @param {Object|Number} attach ID or file
     @return {Object} The attachment for specific attachment key
     */
    getAttachment(attach) {
        let attachData = typeof attach === 'number' ? this.#fileData.fieldsData.attachments[attach] : attach;
        let fieldProperty = this.#fileData.propertyData[attachData.dataId];
        let fieldTypeMapped = MsgReader.CONST.MSG.FIELD.TYPE_MAPPING[MsgReader.#getFieldType(fieldProperty)];
        let fieldData = MsgReader.#getFieldValue(this.#ds, this.#fileData, fieldProperty, fieldTypeMapped);

        return {fileName: attachData.AttachLongFileName, content: fieldData};
    }

    getAttachments() {
        let attachments=[];

        if (this.#fileData.fieldsData && this.#fileData.fieldsData.attachments && this.#fileData.fieldsData.attachments.length > 0) {
            for (const atm of this.#fileData.fieldsData.attachments) {
                let attachment = this.getAttachment(atm);

                    attachments.push({
                        filename: attachment.fileName,
                        contentType: atm.AttachMimeTag,
                        content: attachment.content,
                        filesize: atm.contentLength,
                        pidContentId: atm.AttachmentContentId
                    });

            }
        }

        return attachments;
    }

    getDate() {
        let date = this.getHeader('date');
        if (date) {
            return new Date(date);
        }

        if (this.#fileData.fieldsData && this.#fileData.fieldsData._properties) {
            const props = this.#fileData.fieldsData._properties;

            if (props.ClientSubmitTime && props.ClientSubmitTime.data > new Date(2000,0,1)) {
                return props.ClientSubmitTime.data;
            }
            if (props.DeliveryOrRenewTime && props.DeliveryOrRenewTime.data > new Date(2000,0,1)) {
                return props.DeliveryOrRenewTime.data;
            }
            if (props.CreationTime && props.CreationTime.data > new Date(2000,0,1)) {
                return props.CreationTime.data;
            }
            if (props.LastModificationTime && props.LastModificationTime.data > new Date(2000,0,1)) {
                return props.LastModificationTime.data;
            }

        }

        return null;
    }

    getSubject() {
        const hSub = this.getHeader('subject', true, true);
        if (hSub) {
            return hSub;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.Subject) {
            return this.#fileData.fieldsData.Subject;
        }
    }

    getType() {
        if (this.getHeader('received')) {
            return 'received';
        } else {
            return 'sent';
        }
    }

    getFrom() {
        const hSub = this.getHeader('from', true, true);
        if (hSub) {
            return hSub;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.SenderName && this.#fileData.fieldsData.SentRepresentingSmtpAddress) {
            return this.#fileData.fieldsData.SenderName + ' <' + this.#fileData.fieldsData.SentRepresentingSmtpAddress + '>';

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.SenderName && this.#fileData.fieldsData.LastModifierSMTPAddress) {
            return this.#fileData.fieldsData.SenderName + ' <' + this.#fileData.fieldsData.LastModifierSMTPAddress + '>';

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.SentRepresentingSmtpAddress) {
            return this.#fileData.fieldsData.SentRepresentingSmtpAddress;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.LastModifierSMTPAddress) {
            return this.#fileData.fieldsData.LastModifierSMTPAddress;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.SenderName && this.#fileData.fieldsData.LastModifierName && !this.#fileData.fieldsData.LastModifierName.includes('<')) {
            return this.#fileData.fieldsData.SenderName + ' <' + this.#fileData.fieldsData.LastModifierName + '>';

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.SenderName) {
            return this.#fileData.fieldsData.SenderName;
        }

        return '';
    }

    getBcc() {
        const hSub = this.getHeader('bcc', true, true);
        if (hSub) {
            return hSub;

        } else {
            return this.#getRecipients().bcc;
        }
    }

    getCc() {
        const hSub = this.getHeader('cc', true, true);
        if (hSub) {
            return hSub;

        } else {
            return this.#getRecipients().cc;
        }
    }

    getTo() {
        const hSub = this.getHeader('to', true, true);
        if (hSub) {
            return hSub;

        } else {
            return this.#getRecipients().to;
        }
    }

    getReplyTo() {
        return this.getHeader('reply-to', true, true);
    }

    /**
     * returns a header. If a header occurs more than once, a array is returned.
     * @param {String} key
     * @param {Boolean} decode
     * @param {Boolean} removeLineBreaks
     * @returns {String|Array|null}
     */
    getHeader(key, decode=false, removeLineBreaks=false) {
        let val = null;

        if (this.#headers && this.#headers[key.toLowerCase()]) {
            val = this.#headers[key.toLowerCase()];
        }

        if (val && decode) {
            if (typeof val === 'string') {
                val = this.#decodeRfc1342(val);
            } else {
                val = val.map(this.#decodeRfc1342);
            }
        }

        if (val && removeLineBreaks) {
            if (typeof val === 'string') {
                val = val.replace(/\r?\n\s/g, '');
            } else {
                val = val.map((v) => { return v.replace(/\r?\n\s/g, ''); });
            }
        }

        return val;
    }

    getMessageText() {
        let val = this.#fileData.fieldsData.Body;

        // replace nbsp with space
        let re = new RegExp(String.fromCharCode(160), "g");
        val = val.replace(re, " ");

        // replace multiple newlines
        val = val.replace(/\r?\n\s*\r?\n\s*\r?\n/g, "\n\n");

        return val;
    }

    getMessageHtml() {
        return this.#fileData.fieldsData.BodyHtml;
    }

    // ----------------------------
    // PRIVATE STATIC CONSTANTS
    // ----------------------------

    #getRecipients() {
        const response = {to: [], cc: [], bcc: []};
        const displayTo = this.#fileData.fieldsData.DisplayTo ? this.#fileData.fieldsData.DisplayTo.trim() : '';
        const displayCc = this.#fileData.fieldsData.DisplayCc ? this.#fileData.fieldsData.DisplayCc.trim() : '';
        const displayBcc = this.#fileData.fieldsData.DisplayBcc ? this.#fileData.fieldsData.DisplayBcc.trim() : '';

        if (this.#fileData.fieldsData.recipients) {
            this.#fileData.fieldsData.recipients.forEach((recipient) => {
                const mail = (MsgReader.#getEmailIfValid(recipient.EmailAddress) ?? MsgReader.#getEmailIfValid(recipient.SmtpAddress) ?? recipient.DisplayName ?? '').trim();
                const disp = (recipient.DisplayName ?? '').trim();
                let type = 'to';

                if (disp && displayTo.indexOf(disp) !== -1) {
                    type = 'to';
                } else if (disp && displayCc.indexOf(disp) !== -1) {
                    type = 'cc';
                } else if (disp && displayBcc.indexOf(disp) !== -1) {
                    type = 'bcc';
                }

                if (mail && disp) {
                    response[type].push(disp + ' <' + mail + '>');
                } else if (mail) {
                    response[type].push(mail);
                } else if (disp) {
                    response[type].push(disp);
                }
            });
        }

        return {
            to: response.to.join('; '),
            cc: response.cc.join('; '),
            bcc: response.bcc.join('; ')
        };
    }

    // constants
    static get CONST() {
        return {
            FILE_HEADER: MsgReader.#uInt2int([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
            MSG: {
                UNUSED_BLOCK: -1,
                END_OF_CHAIN: -2,

                S_BIG_BLOCK_SIZE: 0x0200,
                S_BIG_BLOCK_MARK: 9,

                L_BIG_BLOCK_SIZE: 0x1000,
                L_BIG_BLOCK_MARK: 12,

                SMALL_BLOCK_SIZE: 0x0040,
                BIG_BLOCK_MIN_DOC_SIZE: 0x1000,
                HEADER: {
                    PROPERTY_START_OFFSET: 0x30,

                    BAT_START_OFFSET: 0x4c,
                    BAT_COUNT_OFFSET: 0x2C,

                    SBAT_START_OFFSET: 0x3C,
                    SBAT_COUNT_OFFSET: 0x40,

                    XBAT_START_OFFSET: 0x44,
                    XBAT_COUNT_OFFSET: 0x48
                },
                PROP: {
                    NO_INDEX: -1,
                    PROPERTY_SIZE: 0x0080,

                    NAME_SIZE_OFFSET: 0x40,
                    MAX_NAME_LENGTH: (/*NAME_SIZE_OFFSET*/0x40 / 2) - 1,
                    TYPE_OFFSET: 0x42,
                    PREVIOUS_PROPERTY_OFFSET: 0x44,
                    NEXT_PROPERTY_OFFSET: 0x48,
                    CHILD_PROPERTY_OFFSET: 0x4C,
                    START_BLOCK_OFFSET: 0x74,
                    SIZE_OFFSET: 0x78,
                    TYPE_ENUM: {
                        DIRECTORY: 1,
                        DOCUMENT: 2,
                        ROOT: 5
                    }
                },
                FIELD: {
                    PREFIX: {
                        ATTACHMENT: '__attach_version1.0',
                        RECIPIENT: '__recip_version1.0',
                        DOCUMENT: '__substg1.'
                    },

                    // example (use fields as needed)
                    NAME_MAPPING: {

                        // email specific
                       // '0037': 'subject',
                       // '0c1a': 'senderName',
                       // '5d02': 'senderEmail',
                       // '1000': 'body',
                       // '1013': 'bodyHTML',
                       // '007d': 'headers',

                        // attachment specific
                        //'3703': 'extension',
                        //'3704': 'fileNameShort',
                        //'3707': 'fileName',
                        //'3712': 'pidContentId',
                       // '370e': 'mimeType',

                        // recipient specific
                        //'3001': 'name',
                        //'39fe': 'email'
                    },

                    MAPI_PROPERTIES: {
                        "0x0807": "PR_EMS_AB_ROOM_CAPACITY",
                        "0x0809": "PR_EMS_AB_ROOM_DESCRIPTION",
                        "0x3004": "Comment",
                        "0x3007": "CreationTime",
                        "0x3008": "LastModificationTime",
                        "0x3905": "DisplayTypeEx",
                        "0x39fe": "SmtpAddress",
                        "0x39ff": "SimpleDisplayName",
                        "0x3a00": "Account",
                        "0x3a06": "GivenName",
                        "0x3a08": "BusinessTelephoneNumber",
                        "0x3a09": "HomeTelephoneNumber",
                        "0x3a0a": "Initials",
                        "0x3a0f": "MhsCommonName",
                        "0x3a11": "Surname",
                        "0x3a16": "CompanyName",
                        "0x3a17": "Title",
                        "0x3a18": "DepartmentName",
                        "0x3a19": "OfficeLocation",
                        "0x3a1b": "Business2TelephoneNumber",
                        "0x3a1c": "MobileTelephoneNumber",
                        "0x3a21": "PagerTelephoneNumber",
                        "0x3a22": "UserCertificate",
                        "0x3a23": "PrimaryFaxNumber",
                        "0x3a26": "Country",
                        "0x3a27": "Locality",
                        "0x3a28": "StateOrProvince",
                        "0x3a29": "StreetAddress",
                        "0x3a2a": "PostalCode",
                        "0x3a2b": "PostOfficeBox",
                        "0x3a2c": "TelexNumber",
                        "0x3a2e": "AssistantTelephoneNumber",
                        "0x3a2f": "Home2TelephoneNumber",
                        "0x3a30": "Assistant",
                        "0x3a40": "SendRichInfo",
                        "0x3a5d": "HomeAddressStreet",
                        "0x3a70": "UserSMimeCertificate",
                        "0x3a71": "SendInternetEncoding",
                        "0x8004": "PR_EMS_AB_FOLDER_PATHNAME",
                        "0x8005": "PR_EMS_AB_MANAGER_T",
                        "0x8006": "HomeMdb",
                        "0x8007": "dispidContactItemData",
                        "0x8008": "MemberOf",
                        "0x8009": "Members",
                        "0x800c": "ManagedBy",
                        "0x800e": "PR_EMS_AB_REPORTS",
                        "0x800f": "ProxyAddresses",
                        "0x8010": "TemplateInfoHelpFileContents",
                        "0x8011": "PR_EMS_AB_TARGET_ADDRESS",
                        "0x8015": "GrantSendOnBehalfTo",
                        "0x8017": "TemplateInfoTemplate",
                        "0x8018": "TemplateInfoScript",
                        "0x8023": "dispidContactCharSet",
                        "0x8024": "PR_EMS_AB_OWNER_BL_O",
                        "0x8025": "dispidAutoLog",
                        "0x8026": "dispidFileUnderList",
                        "0x8028": "dispidABPEmailList",
                        "0x8029": "dispidABPArrayType",
                        "0x802d": "PR_EMS_AB_EXTENSION_ATTRIBUTE_1",
                        "0x802e": "PR_EMS_AB_EXTENSION_ATTRIBUTE_2",
                        "0x802f": "PR_EMS_AB_EXTENSION_ATTRIBUTE_3",
                        "0x8030": "PR_EMS_AB_EXTENSION_ATTRIBUTE_4",
                        "0x8031": "PR_EMS_AB_EXTENSION_ATTRIBUTE_5",
                        "0x8032": "PR_EMS_AB_EXTENSION_ATTRIBUTE_6",
                        "0x8033": "PR_EMS_AB_EXTENSION_ATTRIBUTE_7",
                        "0x8034": "PR_EMS_AB_EXTENSION_ATTRIBUTE_8",
                        "0x8035": "PR_EMS_AB_EXTENSION_ATTRIBUTE_9",
                        "0x8036": "PR_EMS_AB_EXTENSION_ATTRIBUTE_10",
                        "0x8038": "PfContacts",
                        "0x803b": "TemplateInfoHelpFileName",
                        "0x803c": "ObjectDistinguishedName",
                        "0x8040": "dispidBCDisplayDefinition",
                        "0x8041": "dispidBCCardPicture",
                        "0x8045": "dispidWorkAddressStreet",
                        "0x8046": "dispidWorkAddressCity",
                        "0x8047": "dispidWorkAddressState",
                        "0x8048": "TemplateInfoEmailType",
                        "0x8049": "dispidWorkAddressCountry",
                        "0x804a": "dispidWorkAddressPostOfficeBox",
                        "0x804c": "dispidDLChecksum",
                        "0x804e": "dispidAnniversaryEventEID",
                        "0x804f": "dispidContactUserField1",
                        "0x8051": "dispidContactUserField3",
                        "0x8052": "dispidContactUserField4",
                        "0x8053": "dispidDLName",
                        "0x8054": "dispidDLOneOffMembers",
                        "0x8055": "dispidDLMembers",
                        "0x8062": "dispidInstMsg",
                        "0x8064": "dispidDLStream",
                        "0x806a": "PR_EMS_AB_DELIV_CONT_LENGTH",
                        "0x8073": "PR_EMS_AB_DL_MEM_SUBMIT_PERMS_BL_O",
                        "0x8080": "dispidEmail1DisplayName",
                        "0x8082": "dispidEmail1AddrType",
                        "0x8083": "dispidEmail1EmailAddress",
                        "0x8084": "dispidEmail1OriginalDisplayName",
                        "0x8085": "dispidEmail1OriginalEntryID",
                        "0x8090": "dispidEmail2DisplayName",
                        "0x8092": "dispidEmail2AddrType",
                        "0x8093": "dispidEmail2EmailAddress",
                        "0x8094": "dispidEmail2OriginalDisplayName",
                        "0x8095": "dispidEmail2OriginalEntryID",
                        "0x80a0": "dispidEmail3DisplayName",
                        "0x80a2": "dispidEmail3AddrType",
                        "0x80a3": "dispidEmail3EmailAddress",
                        "0x80a4": "dispidEmail3OriginalDisplayName",
                        "0x80a5": "dispidEmail3OriginalEntryID",
                        "0x80b2": "dispidFax1AddrType",
                        "0x80b3": "dispidFax1EmailAddress",
                        "0x80b4": "dispidFax1OriginalDisplayName",
                        "0x80b5": "dispidFax1OriginalEntryID",
                        "0x80b8": "HideDLMembership",
                        "0x80c2": "dispidFax2AddrType",
                        "0x80c3": "dispidFax2EmailAddress",
                        "0x80c4": "dispidFax2OriginalDisplayName",
                        "0x80c5": "dispidFax2OriginalEntryID",
                        "0x80d2": "dispidFax3AddrType",
                        "0x80d3": "dispidFax3EmailAddress",
                        "0x80d4": "dispidFax3OriginalDisplayName",
                        "0x80d5": "dispidFax3OriginalEntryID",
                        "0x80d8": "dispidFreeBusyLocation",
                        "0x80da": "dispidHomeAddressCountryCode",
                        "0x80db": "dispidWorkAddressCountryCode",
                        "0x80dc": "dispidOtherAddressCountryCode",
                        "0x80dd": "dispidAddressCountryCode",
                        "0x80de": "dispidApptBirthdayLocal",
                        "0x80df": "dispidApptAnniversaryLocal",
                        "0x80e0": "dispidIsContactLinked",
                        "0x80e2": "dispidContactLinkedGALEntryID",
                        "0x80e3": "dispidContactLinkSMTPAddressCache",
                        "0x80e5": "dispidContactLinkLinkRejectHistory",
                        "0x80e6": "dispidContactLinkGALLinkState",
                        "0x80e8": "dispidContactLinkGALLinkID",
                        "0x8101": "dispidTaskStatus",
                        "0x8102": "dispidPercentComplete",
                        "0x8103": "dispidTeamTask",
                        "0x8104": "dispidTaskStartDate",
                        "0x8105": "dispidTaskDueDate",
                        "0x8107": "dispidTaskResetReminder",
                        "0x8108": "dispidTaskAccepted",
                        "0x8109": "dispidTaskDeadOccur",
                        "0x810f": "dispidTaskDateCompleted",
                        "0x8110": "dispidTaskActualEffort",
                        "0x8111": "dispidTaskEstimatedEffort",
                        "0x8112": "dispidTaskVersion",
                        "0x8113": "dispidTaskState",
                        "0x8115": "dispidTaskLastUpdate",
                        "0x8116": "dispidTaskRecur",
                        "0x8117": "dispidTaskMyDelegators",
                        "0x8119": "dispidTaskSOC",
                        "0x811a": "dispidTaskHistory",
                        "0x811b": "dispidTaskUpdates",
                        "0x811e": "dispidTaskFCreator",
                        "0x811f": "dispidTaskOwner",
                        "0x8120": "dispidTaskMultRecips",
                        "0x8121": "dispidTaskDelegator",
                        "0x8122": "dispidTaskLastUser",
                        "0x8123": "dispidTaskOrdinal",
                        "0x8124": "dispidTaskNoCompute",
                        "0x8125": "dispidTaskLastDelegate",
                        "0x8126": "dispidTaskFRecur",
                        "0x8127": "dispidTaskRole",
                        "0x8129": "dispidTaskOwnership",
                        "0x812a": "dispidTaskDelegValue",
                        "0x812c": "dispidTaskFFixOffline",
                        "0x8139": "dispidTaskCustomFlags",
                        "0x8170": "AbNetworkAddress",
                        "0x8202": "dispidApptSeqTime",
                        "0x8c57": "PR_EMS_AB_EXTENSION_ATTRIBUTE_11",
                        "0x8c58": "PR_EMS_AB_EXTENSION_ATTRIBUTE_12",
                        "0x8c59": "PR_EMS_AB_EXTENSION_ATTRIBUTE_13",
                        "0x8c60": "PR_EMS_AB_EXTENSION_ATTRIBUTE_14",
                        "0x8c61": "PR_EMS_AB_EXTENSION_ATTRIBUTE_15",
                        "0x8c6a": "Certificate",
                        "0x8c6d": "ObjectGuid",
                        "0x8c8e": "PR_EMS_AB_PHONETIC_GIVEN_NAME",
                        "0x8c8f": "PR_EMS_AB_PHONETIC_SURNAME",
                        "0x8c90": "PR_EMS_AB_PHONETIC_DEPARTMENT_NAME",
                        "0x8c91": "PR_EMS_AB_PHONETIC_COMPANY_NAME",
                        "0x8c92": "PR_EMS_AB_PHONETIC_DISPLAY_NAME",
                        "0x8c94": "PR_EMS_AB_HAB_SHOW_IN_DEPARTMENTS",
                        "0x8c96": "AddressBookRoomContainers",
                        "0x8c97": "PR_EMS_AB_HAB_DEPARTMENT_MEMBERS",
                        "0x8c98": "PR_EMS_AB_HAB_ROOT_DEPARTMENT",
                        "0x8c99": "PR_EMS_AB_HAB_PARENT_DEPARTMENT",
                        "0x8c9a": "PR_EMS_AB_HAB_CHILD_DEPARTMENTS",
                        "0x8c9e": "ThumbnailPhoto",
                        "0x8ca8": "PR_EMS_AB_ORG_UNIT_ROOT_DN",
                        "0x8cac": "PR_EMS_AB_DL_SENDER_HINT_TRANSLATIONS_W",
                        "0x8cc2": "PR_EMS_AB_UM_SPOKEN_NAME",
                        "0x8cd8": "PR_EMS_AB_AUTH_ORIG",
                        "0x8cd9": "PR_EMS_AB_UNAUTH_ORIG",
                        "0x8cda": "PR_EMS_AB_DL_MEM_SUBMIT_PERMS",
                        "0x8cdb": "PR_EMS_AB_DL_MEM_REJECT_PERMS",
                        "0x8cdd": "PR_EMS_AB_HAB_IS_HIERARCHICAL_GROUP",
                        "0x8ce2": "PR_EMS_AB_DL_TOTAL_MEMBER_COUNT",
                        "0x8ce3": "PR_EMS_AB_DL_EXTERNAL_MEMBER_COUNT",
                        "0x850e": "dispidAgingDontAgeMe",
                        "0x8238": "dispidAllAttendeesString",
                        "0x8246": "dispidAllowExternCheck",
                        "0x8207": "dispidApptAuxFlags",
                        "0x8214": "dispidApptColor",
                        "0x8257": "dispidApptCounterProposal",
                        "0x8213": "dispidApptDuration",
                        "0x8211": "dispidApptEndDate",
                        "0x8210": "dispidApptEndTime",
                        "0x820e": "dispidApptEndWhole",
                        "0x8203": "dispidApptLastSequence",
                        "0x0024": "OriginatorReturnAddress",
                        "0x825a": "dispidApptNotAllowPropose",
                        "0x8259": "dispidApptProposalNum",
                        "0x8256": "dispidApptProposedDuration",
                        "0x8251": "dispidApptProposedEndWhole",
                        "0x8250": "dispidApptProposedStartWhole",
                        "0x8216": "dispidApptRecur",
                        "0x8230": "dispidApptReplyName",
                        "0x8220": "dispidApptReplyTime",
                        "0x8201": "dispidApptSequence",
                        "0x8212": "dispidApptStartDate",
                        "0x820f": "dispidApptStartTime",
                        "0x820d": "dispidApptStartWhole",
                        "0x8217": "dispidApptStateFlags",
                        "0x8215": "dispidApptSubType",
                        "0x825f": "dispidApptTZDefEndDisplay",
                        "0x8260": "dispidApptTZDefRecur",
                        "0x825e": "dispidApptTZDefStartDisplay",
                        "0x825d": "dispidApptUnsendableRecips",
                        "0x8226": "dispidApptUpdateTime",
                        "0x0001": "AcknowledgementMode",
                        "0x823a": "dispidAutoFillLocation",
                        "0x851a": "dispidSniffState",
                        "0x8244": "dispidAutoStartCheck",
                        "0x8535": "dispidBilling",
                        "0x804d": "dispidBirthdayEventEID",
                        "0x8205": "dispidBusyStatus",
                        "0x001c": "LID_CALENDAR_TYPE",
                        "0x9000": "dispidCategories",
                        "0x823c": "dispidCCAttendeesString",
                        "0x8204": "dispidChangeHighlight",
                        "0x85b6": "dispidClassification",
                        "0x85b7": "dispidClassDesc",
                        "0x85b8": "dispidClassGuid",
                        "0x85ba": "dispidClassKeep",
                        "0x85b5": "dispidClassified",
                        "0x0023": "OriginatorDeliveryReportRequested",
                        "0x0015": "ExpiryTime",
                        "0x8236": "dispidClipEnd",
                        "0x8235": "dispidClipStart",
                        "0x8247": "dispidCollaborateDoc",
                        "0x8517": "dispidCommonEnd",
                        "0x8516": "dispidCommonStart",
                        "0x8539": "dispidCompanies",
                        "0x8240": "dispidConfCheck",
                        "0x8241": "dispidConfType",
                        "0x8585": "dispidContactLinkEntry",
                        "0x8586": "dispidContactLinkName",
                        "0x8584": "dispidContactLinkSearchKey",
                        "0x853a": "dispidContacts",
                        "0x8050": "dispidContactUserField2",
                        "0x85ca": "dispidConvActionLastAppliedTime",
                        "0x85c8": "dispidConvActionMaxDeliveryTime",
                        "0x85c6": "dispidConvActionMoveFolderEid",
                        "0x85c7": "dispidConvActionMoveStoreEid",
                        "0x85cb": "dispidConvActionVersion",
                        "0x85c9": "dispidConvExLegacyProcessedRand",
                        "0x8552": "dispidCurrentVersion",
                        "0x8554": "dispidCurrentVersionName",
                        "0x0011": "DiscardReason",
                        "0x1000": "Body",
                        "0x0009": "ContentLength",
                        "0x8242": "dispidDirectory",
                        "0x000f": "DeferredDeliveryTime",
                        "0x0010": "DeliverTime",
                        "0x8228": "dispidExceptionReplaceTime",
                        "0x822b": "dispidFExceptionalAttendees",
                        "0x8206": "dispidFExceptionalBody",
                        "0x8229": "dispidFInvited",
                        "0x8530": "dispidRequest",
                        "0x85c0": "dispidFlagStringEnum",
                        "0x820a": "dispidFwrdInstance",
                        "0x8261": "dispidForwardNotificationRecipients",
                        "0x822f": "dispidFOthersAppt",
                        "0x0003": "AuthorizingUsers",
                        "0x801a": "dispidHomeAddress",
                        "0x802b": "dispidHTML",
                        "0x1001": "ReportText",
                        "0x827a": "InboundICalStream",
                        "0x8224": "dispidIntendedBusyStatus",
                        "0x8580": "dispidInetAcctName",
                        "0x8581": "dispidInetAcctStamp",
                        "0x000a": "ContentReturnRequested",
                        "0x0005": "AutoForwarded",
                        "0x0004": "AutoForwardComment",
                        "0x820c": "dispidLinkedTaskItems",
                        "0x8208": "dispidLocation",
                        "0x8711": "dispidLogDocPosted",
                        "0x870e": "dispidLogDocPrinted",
                        "0x8710": "dispidLogDocRouted",
                        "0x870f": "dispidLogDocSaved",
                        "0x8707": "dispidLogDuration",
                        "0x8708": "dispidLogEnd",
                        "0x870c": "dispidLogFlags",
                        "0x8706": "dispidLogStart",
                        "0x8700": "dispidLogType",
                        "0x8712": "dispidLogTypeDesc",
                        "0x0026": "Priority",
                        "0x8209": "dispidMWSURL",
                        "0x0013": "DlExpansionHistory",
                        "0x1006": "RtfSyncBodyCrc",
                        "0x0017": "Importance",
                        "0x8248": "dispidNetShowURL",
                        "0x100b": "IsIntegJobProgress",
                        "0x8538": "dispidNonSendableBCC",
                        "0x8537": "dispidNonSendableCC",
                        "0x8536": "dispidNonSendableTo",
                        "0x8545": "dispidNonSendBccTrackStatus",
                        "0x8544": "dispidNonSendCcTrackStatus",
                        "0x8543": "dispidNonSendToTrackStatus",
                        "0x8b00": "dispidNoteColor",
                        "0x8b03": "dispidNoteHeight",
                        "0x8b02": "dispidNoteWidth",
                        "0x8b04": "dispidNoteX",
                        "0x8b05": "dispidNoteY",
                        "0x1005": "IsIntegJobCreationTime",
                        "0x0028": "ProofOfSubmissionRequested",
                        "0x0018": "IpmId",
                        "0x002a": "ReceiptTime",
                        "0x0029": "ReadReceiptRequested",
                        "0x8249": "dispidOnlinePassword",
                        "0x0007": "ContentCorrelator",
                        "0x8243": "dispidOrgAlias",
                        "0x8237": "dispidOrigStoreEid",
                        "0x801c": "dispidOtherAddress",
                        "0x001a": "MessageClass",
                        "0x822e": "dispidOwnerName",
                        "0x85e0": "dispidPendingStateforTMDocument",
                        "0x8022": "dispidPostalAddressId",
                        "0x8904": "dispidPostRssChannel",
                        "0x8900": "dispidPostRssChannelLink",
                        "0x8903": "dispidPostRssItemGuid",
                        "0x8902": "dispidPostRssItemHash",
                        "0x8901": "dispidPostRssItemLink",
                        "0x8905": "dispidPostRssItemXml",
                        "0x8906": "dispidPostRssSubscription",
                        "0x8506": "dispidPrivate",
                        "0x100d": "IsIntegJobSource",
                        "0x8232": "dispidRecurPattern",
                        "0x8231": "dispidRecurType",
                        "0x8223": "dispidRecurring",
                        "0x85bd": "dispidReferenceEID",
                        "0x8501": "dispidReminderDelta",
                        "0x851f": "dispidReminderFileParam",
                        "0x851c": "dispidReminderOverride",
                        "0x851e": "dispidReminderPlaySound",
                        "0x8503": "dispidReminderSet",
                        "0x8560": "dispidReminderNextTime",
                        "0x8502": "dispidReminderTime",
                        "0x8505": "dispidReminderTimeDate",
                        "0x8504": "dispidReminderTimeTime",
                        "0x851d": "dispidReminderType",
                        "0x8511": "dispidRemoteStatus",
                        "0x0006": "ContentConfidentialityAlgorithmId",
                        "0x0008": "ContentIdentifier",
                        "0x8218": "dispidResponseStatus",
                        "0x85cc": "dispidExchangeProcessed",
                        "0x85cd": "dispidExchangeProcessingAction",
                        "0x8a19": "dispidSharingAnonymity",
                        "0x8a2d": "dispidSharingBindingEid",
                        "0x8a51": "dispidSharingBrowseUrl",
                        "0x8a17": "dispidSharingCaps",
                        "0x8a24": "dispidSharingConfigUrl",
                        "0x8a45": "dispidSharingDataRangeEnd",
                        "0x8a44": "dispidSharingDataRangeStart",
                        "0x8a2b": "dispidSharingDetail",
                        "0x8a21": "dispidSharingExtXml",
                        "0x8a13": "dispidSharingFilter",
                        "0x8a0a": "dispidSharingFlags",
                        "0x8a18": "dispidSharingFlavor",
                        "0x8a15": "dispidSharingFolderEid",
                        "0x8a2e": "dispidSharingIndexEid",
                        "0x8a09": "dispidSharingInitiatorEid",
                        "0x8a07": "dispidSharingInitiatorName",
                        "0x8a08": "dispidSharingInitiatorSmtp",
                        "0x8a1c": "dispidSharingInstanceGuid",
                        "0x8a55": "dispidSharingLastAutoSync",
                        "0x8a1f": "dispidSharingLastSync",
                        "0x8a4d": "dispidSharingLocalComment",
                        "0x8a23": "dispidSharingLocalLastMod",
                        "0x8a0f": "dispidSharingLocalName",
                        "0x8a0e": "dispidSharingLocalPath",
                        "0x8a49": "dispidSharingLocalStoreUid",
                        "0x8a14": "dispidSharingLocalType",
                        "0x8a10": "dispidSharingLocalUid",
                        "0x8a29": "dispidSharingOriginalMessageEid",
                        "0x8a5c": "dispidSharingParentBindingEid",
                        "0x8a1e": "dispidSharingParticipants",
                        "0x8a1b": "dispidSharingPermissions",
                        "0x8a0b": "dispidSharingProviderExtension",
                        "0x8a01": "dispidSharingProviderGuid",
                        "0x8a02": "dispidSharingProviderName",
                        "0x8a03": "dispidSharingProviderUrl",
                        "0x8a47": "dispidSharingRangeEnd",
                        "0x8a46": "dispidSharingRangeStart",
                        "0x8a1a": "dispidSharingReciprocation",
                        "0x8a4b": "dispidSharingRemoteByteSize",
                        "0x8a2f": "dispidSharingRemoteComment",
                        "0x8a4c": "dispidSharingRemoteCrc",
                        "0x8a22": "dispidSharingRemoteLastMod",
                        "0x8a4f": "dispidSharingRemoteMsgCount",
                        "0x8a05": "dispidSharingRemoteName",
                        "0x8a0d": "dispidSharingRemotePass",
                        "0x8a04": "dispidSharingRemotePath",
                        "0x8a48": "dispidSharingRemoteStoreUid",
                        "0x8a1d": "dispidSharingRemoteType",
                        "0x8a06": "dispidSharingRemoteUid",
                        "0x8a0c": "dispidSharingRemoteUser",
                        "0x8a5b": "dispidSharingRemoteVersion",
                        "0x8a28": "dispidSharingResponseTime",
                        "0x8a27": "dispidSharingResponseType",
                        "0x8a4e": "dispidSharingRoamLog",
                        "0x8a25": "dispidSharingStart",
                        "0x8a00": "dispidSharingStatus",
                        "0x8a26": "dispidSharingStop",
                        "0x8a60": "dispidSharingSyncFlags",
                        "0x8a2a": "dispidSharingSyncInterval",
                        "0x8a2c": "dispidSharingTimeToLive",
                        "0x8a56": "dispidSharingTimeToLiveAuto",
                        "0x8a42": "dispidSharingWorkingHoursDays",
                        "0x8a41": "dispidSharingWorkingHoursEnd",
                        "0x8a40": "dispidSharingWorkingHoursStart",
                        "0x8a43": "dispidSharingWorkingHoursTZ",
                        "0x8510": "dispidSideEffects",
                        "0x827b": "IsSingleBodyICal",
                        "0x8514": "dispidSmartNoAttach",
                        "0x859c": "dispidSpamOriginalFolder",
                        "0x000d": "ConversionWithLossProhibited",
                        "0x000e": "ConvertedEits",
                        "0x811c": "dispidTaskComplete",
                        "0x8519": "dispidTaskGlobalObjId",
                        "0x8518": "dispidTaskMode",
                        "0x000c": "ConversionEits",
                        "0x8234": "dispidTimeZoneDesc",
                        "0x8233": "dispidTimeZoneStruct",
                        "0x823b": "dispidToAttendeesString",
                        "0x85a0": "dispidToDoOrdinalDate",
                        "0x85a1": "dispidToDoSubOrdinal",
                        "0x85a4": "dispidToDoTitle",
                        "0x8582": "dispidUseTNEF",
                        "0x85bf": "dispidValidFlagStringProof",
                        "0x8524": "dispidVerbResponse",
                        "0x8520": "dispidVerbStream",
                        "0x0012": "DisclosureOfRecipients",
                        "0x0002": "AlternateRecipientAllowed",
                        "0x801b": "dispidWorkAddress",
                        "0x0014": "DlExpansionProhibited",
                        "0x802c": "dispidYomiFirstName",
                        "0x0ff4": "Access",
                        "0x3fe0": "AclTable",
                        "0x0ff7": "AccessLevel",
                        "0x36d8": "AdditionalRenEntryIds",
                        "0x36d9": "AdditionalRenEntryIdsEx",
                        "0xfffd": "AbContainerId",
                        "0x8c93": "PR_EMS_AB_DISPLAY_TYPE_EX",
                        "0x663b": "AddressBookEntryId",
                        "0xfffb": "AbIsMaster",
                        "0x6704": "ClientVersion",
                        "0x674f": "ptagAddrbookMID",
                        "0xfffc": "AbParentEntryId",
                        "0x3002": "AddrType",
                        "0x360c": "Anr",
                        "0x301f": "ArchiveDate",
                        "0x301e": "ArchivePeriod",
                        "0x3018": "ArchiveTag",
                        "0x67aa": "Associated",
                        "0x370f": "AttachAdditionalInfo",
                        "0x3711": "AttachmentContentBase",
                        "0x3712": "AttachmentContentId",
                        "0x3713": "AttachContentLocation",
                        "0x3701": "AttachDataObj",
                        "0x3702": "AttachEncoding",
                        "0x3703": "AttachExtension",
                        "0x3704": "AttachFileName",
                        "0x3714": "AttachFlags",
                        "0x3707": "AttachLongFileName",
                        "0x370d": "AttachLongPathName",
                        "0x7fff": "IsContactPhoto",
                        "0x7ffd": "AttachmentCalendarFlags",
                        "0x7ffe": "AttachmentCalendarHidden",
                        "0x7ffa": "AttachmentCalendarLinkId",
                        "0x3705": "AttachMethod",
                        "0x370e": "AttachMimeTag",
                        "0x0e21": "AttachNum",
                        "0x3708": "AttachPathName",
                        "0x371a": "AttachmentPayloadClass",
                        "0x3719": "AttachmentPayloadProviderGuidString",
                        "0x3709": "AttachRendering",
                        "0x0e20": "AttachSize",
                        "0x370a": "AttachTag",
                        "0x370c": "AttachTransportName",
                        "0x10f4": "AttrHidden",
                        "0x10f6": "AttrReadOnly",
                        "0x3fdf": "AutoResponseSuppress",
                        "0x3a42": "Birthday",
                        "0x1096": "BlockStatus",
                        "0x1015": "BodyContentId",
                        "0x1014": "BodyContentLocation",
                        "0x1013": "BodyHtml",
                        "0x3a24": "BusinessFaxNumber",
                        "0x3a51": "BusinessHomePage",
                        "0x3a02": "CallbackTelephoneNumber",
                        "0x6806": "MailboxMiscFlags",
                        "0x3a1e": "CarTelephoneNumber",
                        "0x10c5": "PR_CDO_RECURRENCEID",
                        "0x65e2": "ChangeKey",
                        "0x67a4": "Cn",
                        "0x3a58": "ChildrensNames",
                        "0x6645": "PromotedProperties",
                        "0x0039": "ClientSubmitTime",
                        "0x66c3": "CodePageId",
                        "0x3a57": "CompanyMainPhoneNumber",
                        "0x3a49": "ComputerNetworkName",
                        "0x3ff0": "BackfillTimeout",
                        "0x3613": "ContainerClass",
                        "0x360f": "ContainerContents",
                        "0x3600": "ContainerFlags",
                        "0x360e": "ContainerHierarchy",
                        "0x3602": "ContentCount",
                        "0x4076": "SpamConfidenceLevel",
                        "0x3603": "ContentUnread",
                        "0x3013": "ConversationId",
                        "0x0071": "ConversationIndex",
                        "0x3016": "ConversationIndexTracking",
                        "0x0070": "ConversationTopic",
                        "0x3ff9": "CreatorEntryId",
                        "0x3ff8": "CreatorName",
                        "0x3a4a": "CustomerId",
                        "0x6647": "DeferredActionMessageBackPatched",
                        "0x6646": "HiddenPromotedProperties",
                        "0x36e5": "DefaultPostMsgClass",
                        "0x6741": "OriginalMessageSvrEId",
                        "0x3feb": "DeferredSendNumber",
                        "0x3fef": "DeferredSendTime",
                        "0x3fec": "DeferredSendUnits",
                        "0x3fe3": "OofHistory",
                        "0x686b": "DelegateFlags",
                        "0x0e01": "DeleteAfterSubmit",
                        "0x670b": "DeletedCountTotal",
                        "0x668f": "DeletedOn",
                        "0x3005": "Depth",
                        "0x0e02": "DisplayBcc",
                        "0x0e03": "DisplayCc",
                        "0x3001": "DisplayName",
                        "0x3a45": "DisplayNamePrefix",
                        "0x0e04": "DisplayTo",
                        "0x3900": "DisplayType",
                        "0x3003": "EmailAddress",
                        "0x0061": "EndDate",
                        "0x0fff": "EntryId",
                        "0x7ffc": "AppointmentExceptionEndTime",
                        "0x7ff9": "SExceptionReplaceTime",
                        "0x7ffb": "AppointmentExceptionStartTime",
                        "0x0e84": "http://schemas.microsoft.com/exchange/ntsecuritydescriptor",
                        "0x3fed": "ExpiryNumber",
                        "0x3fee": "ExpiryUnits",
                        "0x36da": "ExtendedFolderFlags",
                        "0x0e99": "ExtendedRuleActions",
                        "0x0e9a": "ExtendedRuleCondition",
                        "0x0e9b": "ExtendedRuleSizeLimit",
                        "0x6804": "ShutoffQuota",
                        "0x1091": "FlagCompleteTime",
                        "0x1090": "FlagStatus",
                        "0x670e": "PR_FLAT_URL_NAME",
                        "0x3610": "FolderAssociatedContents",
                        "0x6748": "Fid",
                        "0x66a8": "FolderFlags",
                        "0x3601": "FolderType",
                        "0x1095": "FollowupIcon",
                        "0x6869": "OutlookFreeBusyMonthCount",
                        "0x36e4": "FreeBusyEntryIds",
                        "0x6849": "ScheduleInfoRecipientLegacyDn",
                        "0x6848": "AssociatedSearchFolderFlags",
                        "0x6847": "AssociatedSearchFolderTag",
                        "0x6868": "PR_FREEBUSY_RANGE_TIMESTAMP",
                        "0x3a4c": "FtpSite",
                        "0x6846": "AssociatedSearchFolderStorageType",
                        "0x3a4d": "Gender",
                        "0x3a05": "Generation",
                        "0x3a07": "GovernmentIdNumber",
                        "0x0e1b": "Hasattach",
                        "0x3fea": "HasDeferredActionMessage",
                        "0x664a": "HasNamedProperties",
                        "0x663a": "HasRules",
                        "0x663e": "HierarchyChangeNumber",
                        "0x4082": "HierRev",
                        "0x3a43": "Hobbies",
                        "0x3a59": "HomeAddressCity",
                        "0x3a5a": "HomeAddressCountry",
                        "0x3a5b": "HomeAddressPostalCode",
                        "0x3a5e": "HomeAddressPostOfficeBox",
                        "0x3a5c": "HomeAddressStateOrProvince",
                        "0x3a25": "HomeFaxNumber",
                        "0x10c4": "urn:schemas:calendar:dtend",
                        "0x10ca": "urn:schemas:calendar:remindernexttime",
                        "0x10c3": "urn:schemas:calendar:dtstart",
                        "0x1080": "IconIndex",
                        "0x666c": "AttachmentInConflict",
                        "0x3f08": "InitialDetailsPane",
                        "0x1042": "InReplyTo",
                        "0x0ff6": "InstanceKey",
                        "0x674e": "InstanceNum",
                        "0x674d": "InstanceId",
                        "0x3fde": "InternetCPID",
                        "0x5902": "INetMailOverrideFormat",
                        "0x1035": "InternetMessageId",
                        "0x1039": "InternetReferences",
                        "0x36d0": "CalendarFolderEntryId",
                        "0x36d1": "ContactsFolderEntryId",
                        "0x36d7": "DraftsFolderEntryId",
                        "0x36d2": "JournalFolderEntryId",
                        "0x36d3": "NotesFolderEntryId",
                        "0x36d4": "TasksFolderEntryId",
                        "0x3a2d": "IsdnNumber",
                        "0x6103": "PR_JUNK_ADD_RECIPS_TO_SSL",
                        "0x6100": "JunkIncludeContacts",
                        "0x6102": "PR_JUNK_PERMANENTLY_DELETE",
                        "0x6107": "PR_JUNK_PHISHING_ENABLE_LINKS",
                        "0x6101": "JunkThreshold",
                        "0x3a0b": "Keyword",
                        "0x3a0c": "Language",
                        "0x3ffb": "LastModifierEntryId",
                        "0x3ffa": "LastModifierName",
                        "0x1081": "LastVerbExecuted",
                        "0x1082": "LastVerbExecutionTime",
                        "0x1043": "ListHelp",
                        "0x1044": "ListSubscribe",
                        "0x1045": "ListUnsubscribe",
                        "0x6709": "LocalCommitTime",
                        "0x670a": "LocalCommitTimeMax",
                        "0x66a1": "LocaleId",
                        "0x3a0d": "Location",
                        "0x661b": "InternetRFC821From",
                        "0x661c": "MailboxOwnerName",
                        "0x3a4e": "ManagerName",
                        "0x0ff8": "MappingSignature",
                        "0x666d": "SearchAttachments",
                        "0x6671": "MemberId",
                        "0x6672": "MemberName",
                        "0x6673": "MemberRights",
                        "0x0e13": "MessageAttachments",
                        "0x0058": "MessageCcMe",
                        "0x3ffd": "MessageCodePage",
                        "0x0e06": "MessageDeliveryTime",
                        "0x5909": "MessageEditorFormat",
                        "0x0e07": "MessageFlags",
                        "0x3ff1": "MessageLocaleId",
                        "0x0059": "MessageRecipMe",
                        "0x0e12": "MessageRecipients",
                        "0x0e08": "MessageSize",
                        "0x0e17": "MsgStatus",
                        "0x0047": "MessageSubmissionId",
                        "0x0057": "MessageToMe",
                        "0x674a": "Mid",
                        "0x3a44": "MiddleName",
                        "0x64f0": "MimeSkeleton",
                        "0x1016": "NativeBodyInfo",
                        "0x0e29": "NextSendAccount",
                        "0x3a4f": "Nickname",
                        "0x0c05": "NdrDiagCode",
                        "0x0c04": "NdrReasonCode",
                        "0x0c20": "NDRStatusCode",
                        "0x0c06": "NonReceiptNotificationRequested",
                        "0x0e1d": "NormalizedSubject",
                        "0x0ffe": "ObjectType",
                        "0x6802": "SenderTelephoneNumber",
                        "0x6803": "SendOutlookRecallReport",
                        "0x6800": "PR_OAB_NAME",
                        "0x6801": "VoiceMessageDuration",
                        "0x6805": "VoiceMessageAttachmentOrder",
                        "0x36e2": "PR_ORDINAL_MOST",
                        "0x3a10": "OrganizationalIdNumber",
                        "0x004c": "OriginalAuthorEntryId",
                        "0x004d": "OriginalAuthorName",
                        "0x0055": "OriginalDeliveryTime",
                        "0x0072": "OriginalDisplayBcc",
                        "0x0073": "OriginalDisplayCc",
                        "0x0074": "OriginalDisplayTo",
                        "0x3a12": "OriginalEntryId",
                        "0x004b": "OrigMessageClass",
                        "0x1046": "OriginalInternetMessageId",
                        "0x0066": "OriginalSenderAddrType",
                        "0x0067": "OriginalSenderEmailAddress",
                        "0x005b": "OriginalSenderEntryId",
                        "0x005a": "OriginalSenderName",
                        "0x005c": "OriginalSenderSearchKey",
                        "0x002e": "OriginalSensitivity",
                        "0x0068": "OriginalSentRepresentingAddrType",
                        "0x0069": "OriginalSentRepresentingEmailAddress",
                        "0x005e": "OriginalSentRepresentingEntryId",
                        "0x005d": "OriginalSentRepresentingName",
                        "0x005f": "OriginalSentRepresentingSearchKey",
                        "0x0049": "OriginalSubject",
                        "0x004e": "OriginalSubmitTime",
                        "0x0c08": "OriginatorNonDeliveryReportRequested",
                        "0x7c24": "OscSyncEnabledOnServer",
                        "0x3a5f": "OtherAddressCity",
                        "0x3a60": "OtherAddressCountry",
                        "0x3a61": "OtherAddressPostalCode",
                        "0x3a64": "OtherAddressPostOfficeBox",
                        "0x3a62": "OtherAddressStateOrProvince",
                        "0x3a63": "OtherAddressStreet",
                        "0x3a1f": "OtherTelephoneNumber",
                        "0x661d": "OofState",
                        "0x0062": "OwnerApptId",
                        "0x0e09": "ParentEntryId",
                        "0x6749": "ParentFid",
                        "0x0025": "ParentKey",
                        "0x65e1": "ParentSourceKey",
                        "0x3a50": "PersonalHomePage",
                        "0x3019": "PolicyTag",
                        "0x3a15": "PostalAddress",
                        "0x65e3": "PredecessorChangeList",
                        "0x0e28": "PrimarySendAccount",
                        "0x3a1a": "PrimaryTelephoneNumber",
                        "0x7d01": "IsProcessed",
                        "0x3a46": "Profession",
                        "0x666a": "ProhibitReceiveQuota",
                        "0x666e": "ProhibitSendQuota",
                        "0x4083": "PurportedSenderDomain",
                        "0x3a1d": "RadioTelephoneNumber",
                        "0x0e69": "Read",
                        "0x4029": "ReadReceiptAddrType",
                        "0x402a": "ReadReceiptEmailAddress",
                        "0x0046": "ReadReceiptEntryId",
                        "0x402b": "ReadReceiptDisplayName",
                        "0x0053": "ReadReceiptSearchKey",
                        "0x5d05": "ReadReceiptSMTPAddress",
                        "0x0075": "ReceivedByAddrType",
                        "0x0076": "ReceivedByEmailAddress",
                        "0x003f": "ReceivedByEntryId",
                        "0x0040": "ReceivedByName",
                        "0x0051": "ReceivedBySearchKey",
                        "0x5d07": "ReceivedBySmtpAddress",
                        "0x0077": "RcvdRepresentingAddrType",
                        "0x0078": "RcvdRepresentingEmailAddress",
                        "0x0043": "RcvdRepresentingEntryId",
                        "0x0044": "RcvdRepresentingName",
                        "0x0052": "RcvdRepresentingSearchKey",
                        "0x5d08": "RcvdRepresentingSmtpAddress",
                        "0x5ff6": "RecipientDisplayName",
                        "0x5ff7": "RecipientEntryId",
                        "0x5ffd": "RecipientFlags",
                        "0x5fdf": "RecipientOrder",
                        "0x5fe1": "RecipientProposed",
                        "0x5fe4": "RecipientProposedEndTime",
                        "0x5fe3": "RecipientProposedStartTime",
                        "0x002b": "RecipientReassignmentProhibited",
                        "0x5fff": "RecipientTrackStatus",
                        "0x5ffb": "RecipientTrackStatusTime",
                        "0x0c15": "RecipientType",
                        "0x0ff9": "RecordKey",
                        "0x3a47": "PreferredByName",
                        "0x36d5": "RemindersSearchFolderEntryId",
                        "0x0c21": "RemoteMta",
                        "0x370b": "RenderingPosition",
                        "0x004f": "ReplyRecipientEntries",
                        "0x0050": "ReplyRecipientNames",
                        "0x0c17": "ReplyRequested",
                        "0x65c2": "ReplyTemplateID",
                        "0x0030": "ReplyTime",
                        "0x0080": "ReportDisposition",
                        "0x0081": "ReportDispositionMode",
                        "0x0045": "ReportEntryId",
                        "0x6820": "ReportingMta",
                        "0x003a": "ReportName",
                        "0x0054": "ReportSearchKey",
                        "0x0031": "ReportTag",
                        "0x0032": "ReportTime",
                        "0x3fe7": "ResolveMethod",
                        "0x0063": "ResponseRequested",
                        "0x0e0f": "Responsibility",
                        "0x301c": "RetentionDate",
                        "0x301d": "RetentionFlags",
                        "0x301a": "RetentionPeriod",
                        "0x6639": "AccessRights",
                        "0x7c06": "UserConfigurationType",
                        "0x7c07": "UserConfigurationDictionary",
                        "0x7c08": "UserConfigurationXml",
                        "0x3000": "RowId",
                        "0x0ff5": "RowType",
                        "0x1009": "RtfCompressed",
                        "0x0e1f": "RtfInSync",
                        "0x6650": "RuleActionNumber",
                        "0x6680": "RuleActions",
                        "0x6649": "RuleActionType",
                        "0x6679": "RuleCondition",
                        "0x6648": "RuleError",
                        "0x6651": "RuleFolderEntryID",
                        "0x6674": "RuleID",
                        "0x6675": "RuleIDs",
                        "0x6683": "RuleLevel",
                        "0x65ed": "RuleMsgLevel",
                        "0x65ec": "RuleMsgName",
                        "0x65eb": "RuleMsgProvider",
                        "0x65ee": "RuleMsgProviderData",
                        "0x65f3": "RuleMsgSequence",
                        "0x65e9": "RuleMsgState",
                        "0x65ea": "RuleMsgUserFlags",
                        "0x6682": "RuleName",
                        "0x6681": "RuleProvider",
                        "0x6684": "RuleProviderData",
                        "0x6676": "RuleSequence",
                        "0x6677": "RuleState",
                        "0x6678": "RuleUserFlags",
                        "0x686a": "AppointmentTombstonesId",
                        "0x686d": "PR_SCHDINFO_AUTO_ACCEPT_APPTS",
                        "0x6845": "DelegateEntryIds",
                        "0x6844": "ActivityLogicalItemId",
                        "0x684a": "DelegateNames",
                        "0x6842": "DelegateBossWantsCopy",
                        "0x684b": "DelegateBossWantsInfo",
                        "0x686f": "PR_SCHDINFO_DISALLOW_OVERLAPPING_APPTS",
                        "0x686e": "PR_SCHDINFO_DISALLOW_RECURRING_APPTS",
                        "0x6843": "ActivityContainerType",
                        "0x686c": "PR_SCHDINFO_FREEBUSY",
                        "0x6856": "AgingFileName9AndPrev",
                        "0x6854": "NavigationNodeAddressBookEntryId",
                        "0x6850": "ScheduleInfoFreeBusyMerged",
                        "0x6852": "NavigationNodeGroupSection",
                        "0x6855": "ScheduleInfoMonthsOof",
                        "0x6853": "NavigationNodeCalendarColor",
                        "0x684f": "ScheduleInfoMonthsMerged",
                        "0x6851": "NavigationNodeGroupName",
                        "0x6841": "AssociatedSearchFolderTemplateId",
                        "0x6622": "FreeBusyEntryId",
                        "0x683a": "ActivitySequenceNumber",
                        "0x6834": "AssociatedSearchFolderLastUsedTime",
                        "0x300b": "SearchKey",
                        "0x0e6a": "NTSDAsXML",
                        "0x3609": "Selectable",
                        "0x0c1e": "SenderAddrType",
                        "0x0c1f": "SenderEmailAddress",
                        "0x0c19": "SenderEntryId",
                        "0x4079": "SenderIdStatus",
                        "0x0c1a": "SenderName",
                        "0x0c1d": "SenderSearchKey",
                        "0x5d01": "SenderSmtpAddress",
                        "0x0036": "Sensitivity",
                        "0x6740": "SentMailSvrEId",
                        "0x0064": "SentRepresentingAddrType",
                        "0x0065": "SentRepresentingEmailAddress",
                        "0x0041": "SentRepresentingEntryId",
                        "0x401a": "SentRepresentingFlags",
                        "0x0042": "SentRepresentingName",
                        "0x003b": "SentRepresentingSearchKey",
                        "0x5d02": "SentRepresentingSmtpAddress",
                        "0x6638": "FolderChildCount32",
                        "0x6705": "SortLocaleId",
                        "0x65e0": "SourceKey",
                        "0x3a48": "SpouseName",
                        "0x0060": "StartDate",
                        "0x301b": "StartDateEtc",
                        "0x0ffb": "StoreEntryid",
                        "0x340e": "StoreState",
                        "0x340d": "StoreSupportMask",
                        "0x360a": "SubFolders",
                        "0x0037": "Subject",
                        "0x003d": "SubjectPrefix",
                        "0x0c1b": "SupplementaryInfo",
                        "0x0e2d": "SwappedToDoData",
                        "0x0e2c": "SwappedToDoStore",
                        "0x3010": "TargetEntryId",
                        "0x3a4b": "TtytddPhoneNumber",
                        "0x3902": "TemplateId",
                        "0x371b": "TextAttachmentCharset",
                        "0x007f": "TnefCorrelationKey",
                        "0x0e2b": "ToDoItemFlags",
                        "0x3a20": "TransmitableDisplayName",
                        "0x007d": "TransportMessageHeaders",
                        "0x0e79": "TrustSender",
                        "0x6619": "InternetAddressConversion",
                        "0x7001": "UserInformationLastMaintenanceTime",
                        "0x7006": "PR_VD_NAME",
                        "0x7002": "PR_VD_STRINGS",
                        "0x7007": "PR_VD_VERSION",
                        "0x3a41": "WeddingAnniversary",
                        "0x6891": "ConversationContentUnreadMailboxWide",
                        "0x6890": "ConversationContentUnread",
                        "0x684c": "NavigationNodeEntryId",
                        "0x684d": "NavigationNodeRecordKey",
                        "0x6892": "ConversationMessageSize",
                        "0x684e": "NavigationNodeStoreEntryId",
                        "0x0000": "Null",
                        "0x6700": "PstPath",
                        "0x6603": "ProfileUser",
                        "0x660b": "ProfileMailbox",
                        "0x6602": "ProfileHomeServer",
                        "0x6612": "ProfileHomeServerDn",
                        "0x660c": "ProfileServer",
                        "0x6614": "ProfileServerDn",
                        "0x6601": "ProfileConfigFlags",
                        "0x6605": "ProfileTransportFlags",
                        "0x6600": "ProfileVersion",
                        "0x6604": "ProfileConnectFlags",
                        "0x6606": "ProfileUiState",
                        "0x000b": "ConversationKey",
                        "0x0016": "ImplicitConversionProhibited",
                        "0x0019": "LatestDeliveryTime",
                        "0x001b": "MessageDeliveryId",
                        "0x001e": "MessageSecurityLabel",
                        "0x001f": "ObsoletedIpms",
                        "0x0020": "OriginallyIntendedRecipientName",
                        "0x0021": "OriginalEits",
                        "0x0022": "OriginatorCertificate",
                        "0x0027": "OriginCheck",
                        "0x002c": "RedirectionHistory",
                        "0x002d": "RelatedIpms",
                        "0x002f": "Languages",
                        "0x0033": "ReturnedIpm",
                        "0x0034": "Security",
                        "0x0035": "IncompleteCopy",
                        "0x0038": "SubjectIpm",
                        "0x003c": "X400ContentType",
                        "0x003e": "NonReceiptReason",
                        "0x0048": "ProviderSubmitTime",
                        "0x004a": "DiscVal",
                        "0x0056": "OriginalAuthorSearchKey",
                        "0x0079": "OriginalAuthorAddrType",
                        "0x007a": "OriginalAuthorEmailAddress",
                        "0x007b": "OriginallyIntendedRecipAddrType",
                        "0x007c": "OriginallyIntendedRecipEmailAddress",
                        "0x007e": "Delegation",
                        "0x4031": "SentRepresentingSimpleDisplayName",
                        "0x4032": "OriginalSenderSimpleDisplayName",
                        "0x4033": "OriginalSentRepresentingSimpleDisplayName",
                        "0x4034": "ReceivedBySimpleDisplayName",
                        "0x4035": "RcvdRepresentingSimpleDisplayName",
                        "0x4036": "ReadReceiptSimpleDisplayName",
                        "0x4060": "OriginalAuthorSimpleDisplayName",
                        "0x1002": "OriginatorAndDlExpansionHistory",
                        "0x1003": "ReportingDlName",
                        "0x1004": "ReportingMtaCertificate",
                        "0x1007": "RtfSyncBodyCount",
                        "0x1008": "RtfSyncBodyTag",
                        "0x1010": "RtfSyncPrefixCount",
                        "0x1011": "RtfSyncTrailingCount",
                        "0x1012": "OriginallyIntendedRecipEntryId",
                        "0x1097": "ItemTemporaryFlag",
                        "0x10c6": "DavSubmitData",
                        "0x0c00": "ContentIntegrityCheck",
                        "0x0c01": "ExplicitConversion",
                        "0x0c02": "IpmReturnRequested",
                        "0x0c03": "MessageToken",
                        "0x0c07": "DeliveryPoint",
                        "0x0c09": "OriginatorRequestedAlternateRecipient",
                        "0x0c0a": "PhysicalDeliveryBureauFaxDelivery",
                        "0x0c0b": "PhysicalDeliveryMode",
                        "0x0c0c": "PhysicalDeliveryReportRequest",
                        "0x0c0d": "PhysicalForwardingAddress",
                        "0x0c0e": "PhysicalForwardingAddressRequested",
                        "0x0c0f": "PhysicalForwardingProhibited",
                        "0x0c10": "PhysicalRenditionAttributes",
                        "0x0c11": "ProofOfDelivery",
                        "0x0c12": "ProofOfDeliveryRequested",
                        "0x0c13": "RecipientCertificate",
                        "0x0c14": "RecipientNumberForAdvice",
                        "0x0c16": "RegisteredMailType",
                        "0x0c18": "RequestedDeliveryMethod",
                        "0x0c1c": "TypeOfMtsUser",
                        "0x5903": "INetMailOverrideCharset",
                        "0x4030": "SenderSimpleDisplayName",
                        "0x0e00": "CurrentVersion",
                        "0x0e05": "ParentDisplay",
                        "0x0e0a": "SentMailEntryId",
                        "0x0e0c": "Correlate",
                        "0x0e0d": "CorrelateMtsid",
                        "0x0e0e": "DiscreteValues",
                        "0x0e10": "SpoolerStatus",
                        "0x0e11": "TransportStatus",
                        "0x0e14": "SubmitFlags",
                        "0x0e15": "RecipientStatus",
                        "0x0e16": "TransportKey",
                        "0x0e18": "MessageDownloadTime",
                        "0x0e19": "CreationVersion",
                        "0x0e1a": "ModifyVersion",
                        "0x0e1c": "BodyCrc",
                        "0x0e22": "Preprocess",
                        "0x0e23": "InternetArticleNumber",
                        "0x0e25": "OriginatingMtaCertificate",
                        "0x0e26": "ProofOfSubmission",
                        "0x0e3a": "MessageDeepAttachments",
                        "0x0e85": "AntiVirusVendor",
                        "0x0e86": "AntiVirusVersion",
                        "0x0e87": "AntiVirusScanStatus",
                        "0x0e88": "AntiVirusScanInfo",
                        "0x0e96": "TransportAntiVirusStamp",
                        "0x0e7c": "MessageDatabasePage",
                        "0x0e7d": "MessageHeaderDatabasePage",
                        "0x0ffd": "Icon",
                        "0x0ffc": "MiniIcon",
                        "0x0ffa": "StoreRecordKey",
                        "0x0f0e": "MessageAnnotation",
                        "0x3006": "ProviderDisplay",
                        "0x3009": "ResourceFlags",
                        "0x300a": "ProviderDllName",
                        "0x300c": "ProviderUid",
                        "0x300d": "ProviderOrdinal",
                        "0x3012": "ConversationIdObsolete",
                        "0x3014": "BodyTag",
                        "0x3015": "ConversationIndexTrackingObsolete",
                        "0x6827": "ConversationIdHash",
                        "0x6200": "InternetMessageIdHash",
                        "0x6201": "ConversationTopicHash",
                        "0x3660": "ConversationTopicHashEntries",
                        "0x3301": "FormVersion",
                        "0x3302": "FormClsid",
                        "0x3303": "FormContactName",
                        "0x3304": "FormCategory",
                        "0x3305": "FormCategorySub",
                        "0x3306": "FormHostMap",
                        "0x3307": "FormHidden",
                        "0x3308": "FormDesignerName",
                        "0x3309": "FormDesignerGuid",
                        "0x330a": "FormMessageBehavior",
                        "0x3400": "DefaultStore",
                        "0x3410": "IpmSubtreeSearchKey",
                        "0x3411": "IpmOutboxSearchKey",
                        "0x3412": "IpmWastebasketSearchKey",
                        "0x3413": "IpmSentMailSearchKey",
                        "0x3414": "MdbProvider",
                        "0x3415": "ReceiveFolderSettings",
                        "0x35df": "ValidFolderMask",
                        "0x35e0": "IpmSubtreeEntryId",
                        "0x35e2": "IpmOutboxEntryId",
                        "0x35e4": "IpmSentMailEntryId",
                        "0x35e5": "ViewsEntryId",
                        "0x35e6": "CommonViewsEntryId",
                        "0x35e7": "FinderEntryId",
                        "0x35ec": "ConversationsFolderEntryId",
                        "0x35ee": "AllItemsEntryId",
                        "0x35ef": "SharingFolderEntryId",
                        "0x0e5c": "CISearchEnabled",
                        "0x3416": "LocalDirectoryEntryId",
                        "0x3646": "OwnerLogonUserConfigurationCache",
                        "0x3420": "ControlDataForCalendarRepairAssistant",
                        "0x3421": "ControlDataForSharingPolicyAssistant",
                        "0x3422": "ControlDataForElcAssistant",
                        "0x3423": "ControlDataForTopNWordsAssistant",
                        "0x3424": "ControlDataForJunkEmailAssistant",
                        "0x3425": "ControlDataForCalendarSyncAssistant",
                        "0x3426": "ExternalSharingCalendarSubscriptionCount",
                        "0x3427": "ControlDataForUMReportingAssistant",
                        "0x3428": "HasUMReportData",
                        "0x3429": "InternetCalendarSubscriptionCount",
                        "0x342a": "ExternalSharingContactSubscriptionCount",
                        "0x342b": "JunkEmailSafeListDirty",
                        "0x342c": "IsTopNEnabled",
                        "0x342d": "LastSharingPolicyAppliedId",
                        "0x342e": "LastSharingPolicyAppliedHash",
                        "0x342f": "LastSharingPolicyAppliedTime",
                        "0x3430": "OofScheduleStart",
                        "0x3431": "OofScheduleEnd",
                        "0x35fe": "UnsearchableItemsStream",
                        "0x3604": "CreateTemplates",
                        "0x3605": "DetailsTable",
                        "0x3607": "Search",
                        "0x360b": "Status",
                        "0x360d": "ContentsSortOrder",
                        "0x3611": "DefCreateDl",
                        "0x3612": "DefCreateMailuser",
                        "0x3614": "ContainerModifyVersion",
                        "0x3615": "AbProviderId",
                        "0x3616": "DefaultViewEntryId",
                        "0x3617": "AssocContentCount",
                        "0x3644": "SearchFolderMsgCount",
                        "0x361f": "AllowAgeout",
                        "0x6787": "SearchBacklinkNames",
                        "0x3700": "AttachmentX400Parameters",
                        "0x3716": "AttachDisposition",
                        "0x0eb0": "SearchResultKind",
                        "0x0eab": "SearchFullText",
                        "0x0eac": "SearchFullTextSubject",
                        "0x0ead": "SearchFullTextBody",
                        "0x0eae": "SearchFullTextConversationIndex",
                        "0x0eaf": "SearchAllIndexedProps",
                        "0x0eb1": "SearchRecipients",
                        "0x0eb2": "SearchRecipientsTo",
                        "0x0eb3": "SearchRecipientsCc",
                        "0x0eb4": "SearchRecipientsBcc",
                        "0x0eb5": "SearchAccountTo",
                        "0x0eb6": "SearchAccountCc",
                        "0x0eb7": "SearchAccountBcc",
                        "0x0eb8": "SearchEmailAddressTo",
                        "0x0eb9": "SearchEmailAddressCc",
                        "0x0eba": "SearchEmailAddressBcc",
                        "0x0ebb": "SearchSmtpAddressTo",
                        "0x0ebc": "SearchSmtpAddressCc",
                        "0x0ebd": "SearchSmtpAddressBcc",
                        "0x0ebe": "SearchSender",
                        "0x0ebf": "SendYearHigh",
                        "0x0ec0": "SendYearLow",
                        "0x0ec1": "SendMonth",
                        "0x0ec2": "SendDayHigh",
                        "0x0ec3": "SendDayLow",
                        "0x0ec4": "SendQuarterHigh",
                        "0x0ec5": "SendQuarterLow",
                        "0x0ec6": "RcvdYearHigh",
                        "0x0ec7": "RcvdYearLow",
                        "0x0ec8": "RcvdMonth",
                        "0x0ec9": "RcvdDayHigh",
                        "0x0eca": "RcvdDayLow",
                        "0x0ecb": "RcvdQuarterHigh",
                        "0x0ecc": "RcvdQuarterLow",
                        "0x0ecd": "IsIrmMessage",
                        "0x3a01": "AlternateRecipient",
                        "0x3a03": "ConversionProhibited",
                        "0x3a04": "DiscloseRecipients",
                        "0x3a0e": "MailPermission",
                        "0x3a13": "OriginalDisplayName",
                        "0x3a14": "OriginalSearchKey",
                        "0x3a52": "ContactVersion",
                        "0x3a53": "ContactEntryIds",
                        "0x3a54": "ContactAddrTypes",
                        "0x3a55": "ContactDefaultAddressIndex",
                        "0x3a56": "ContactEmailAddresses",
                        "0x3d00": "StoreProviders",
                        "0x3d01": "AbProviders",
                        "0x3d02": "TransportProviders",
                        "0x3d04": "DefaultProfile",
                        "0x3d05": "AbSearchPath",
                        "0x3d06": "AbDefaultDir",
                        "0x3d07": "AbDefaultPab",
                        "0x3d08": "FilteringHooks",
                        "0x3d09": "ServiceName",
                        "0x3d0a": "ServiceDllName",
                        "0x3d0c": "ServiceUid",
                        "0x3d0d": "ServiceExtraUids",
                        "0x3d0e": "Services",
                        "0x3d0f": "ServiceSupportFiles",
                        "0x3d10": "ServiceDeleteFiles",
                        "0x3d11": "AbSearchPathUpdate",
                        "0x3d12": "ProfileName",
                        "0x3fe9": "EformsLocaleId",
                        "0x6620": "NonIpmSubtreeEntryId",
                        "0x6621": "EFormsRegistryEntryId",
                        "0x6623": "OfflineAddressBookEntryId",
                        "0x6624": "LocaleEFormsRegistryEntryId",
                        "0x6625": "LocalSiteFreeBusyEntryId",
                        "0x6626": "LocalSiteOfflineAddressBookEntryId",
                        "0x6810": "OofStateEx",
                        "0x6813": "OofStateUserChangeTime",
                        "0x6814": "UserOofSettingsItemId",
                        "0x35e3": "IpmWasteBasketEntryId",
                        "0x3011": "ForceUserClientBackoff",
                        "0x6618": "InTransitStatus",
                        "0x662a": "TransferEnabled",
                        "0x65f6": "ImapSubscribeList",
                        "0x676c": "MapiEntryIdGuid",
                        "0x662f": "FastTransfer",
                        "0x681a": "MailboxQuarantined",
                        "0x6761": "NextLocalId",
                        "0x6670": "LongTermEntryIdFromTable",
                        "0x0e27": "NTSD",
                        "0x3d21": "AdminNTSD",
                        "0x0f00": "FreeBusyNTSD",
                        "0x0e3f": "AclTableAndNTSD",
                        "0x6707": "UrlName",
                        "0x6698": "ReplicaList",
                        "0x663f": "HasModerator",
                        "0x3fe6": "PublishInAddressBook",
                        "0x6699": "OverallAgeLimit",
                        "0x66c4": "RetentionAgeLimit",
                        "0x6779": "PfQuotaStyle",
                        "0x677b": "PfStorageQuota",
                        "0x6721": "PfOverHardQuotaLimit",
                        "0x6722": "PfMsgSizeLimit",
                        "0x6690": "ReplicationStyle",
                        "0x6691": "ReplicationSchedule",
                        "0x66c5": "DisablePeruserRead",
                        "0x6701": "PfMsgAgeLimit",
                        "0x671d": "PfProxy",
                        "0x3d2f": "SystemFolderFlags",
                        "0x3fd9": "Preview",
                        "0x6828": "LocalDirectory",
                        "0x0e63": "SendFlags",
                        "0x3fe1": "RulesTable",
                        "0x4080": "OofReplyType",
                        "0x4081": "ElcAutoCopyLabel",
                        "0x6716": "ElcAutoCopyTag",
                        "0x6717": "ElcMoveDate",
                        "0x6829": "MemberEmail",
                        "0x682a": "MemberExternalId",
                        "0x682b": "MemberSID",
                        "0x662c": "HierarchySynchronizer",
                        "0x662d": "ContentsSynchronizer",
                        "0x662e": "Collector",
                        "0x6609": "SendRichInfoOnly",
                        "0x6637": "SendNativeBody",
                        "0x6608": "InternetTransmitInfo",
                        "0x6610": "InternetMessageFormat",
                        "0x6611": "InternetMessageTextFormat",
                        "0x6615": "InternetRequestLines",
                        "0x6616": "InternetHeaderLength",
                        "0x6617": "InternetAddressingOptions",
                        "0x661a": "InternetTemporaryFilename",
                        "0x6570": "InternetExternalNewsItem",
                        "0x661e": "InternetRequestHeaders",
                        "0x6631": "InternetClientHostIPName",
                        "0x66c0": "ConnectTime",
                        "0x66c1": "ConnectFlags",
                        "0x66c2": "LogonCount",
                        "0x669f": "HostAddress",
                        "0x66a0": "NTUserName",
                        "0x66a2": "LastLogonTime",
                        "0x66a3": "LastLogoffTime",
                        "0x66a4": "StorageLimitInformation",
                        "0x66a5": "InternetMdns",
                        "0x669b": "DeletedMessageSizeExtended",
                        "0x669c": "DeletedNormalMessageSizeExtended",
                        "0x669d": "DeleteAssocMessageSizeExtended",
                        "0x6640": "DeletedMsgCount",
                        "0x6708": "DateDiscoveredAbsentInDS",
                        "0x6778": "AdminNickName",
                        "0x6723": "QuotaReceiveThreshold",
                        "0x66df": "LastOpTime",
                        "0x66ff": "PacketRate",
                        "0x66cc": "LogonTime",
                        "0x66cd": "LogonFlags",
                        "0x6764": "MsgHeaderFid",
                        "0x66c9": "MailboxDisplayName",
                        "0x66c8": "MailboxDN",
                        "0x66cb": "UserDisplayName",
                        "0x66ca": "UserDN",
                        "0x6710": "SessionId",
                        "0x66cf": "OpenMessageCount",
                        "0x66d0": "OpenFolderCount",
                        "0x66d1": "OpenAttachCount",
                        "0x66d2": "OpenContentCount",
                        "0x66d3": "OpenHierarchyCount",
                        "0x66d4": "OpenNotifyCount",
                        "0x66d5": "OpenAttachTableCount",
                        "0x66d6": "OpenACLTableCount",
                        "0x66d7": "OpenRulesTableCount",
                        "0x66d8": "OpenStreamsCount",
                        "0x66d9": "OpenFXSrcStreamCount",
                        "0x66da": "OpenFXDestStreamCount",
                        "0x66db": "OpenContentRegularCount",
                        "0x66dc": "OpenContentCategCount",
                        "0x66dd": "OpenContentRestrictedCount",
                        "0x66de": "OpenContentCategAndRestrictedCount",
                        "0x66b8": "MessagingOpRate",
                        "0x66b9": "FolderOpRate",
                        "0x66ba": "TableOpRate",
                        "0x66bb": "TransferOpRate",
                        "0x66bc": "StreamOpRate",
                        "0x66bd": "ProgressOpRate",
                        "0x66be": "OtherOpRate",
                        "0x66bf": "TotalOpRate",
                        "0x671f": "PfProxyRequired",
                        "0x6724": "ClientIP",
                        "0x6725": "ClientIPMask",
                        "0x6726": "ClientMacAddress",
                        "0x6727": "ClientMachineName",
                        "0x6728": "ClientAdapterSpeed",
                        "0x6729": "ClientRpcsAttempted",
                        "0x672a": "ClientRpcsSucceeded",
                        "0x672b": "ClientRpcErrors",
                        "0x672c": "ClientLatency",
                        "0x666f": "SubmittedByAdmin",
                        "0x680d": "ObjectClassFlags",
                        "0x682d": "MaxMessageSize",
                        "0x672d": "TimeInServer",
                        "0x672e": "TimeInCPU",
                        "0x672f": "ROPCount",
                        "0x6730": "PageRead",
                        "0x6731": "PagePreread",
                        "0x6732": "LogRecordCount",
                        "0x6733": "LogRecordBytes",
                        "0x6734": "LdapReads",
                        "0x6735": "LdapSearches",
                        "0x6736": "DigestCategory",
                        "0x6737": "SampleId",
                        "0x6738": "SampleTime",
                        "0x661f": "DeferredActionFolderEntryID",
                        "0x3fff": "RulesSize",
                        "0x668d": "RuleVersion",
                        "0x65f2": "RuleMsgVersion",
                        "0x3642": "DAMReferenceMessageEntryID",
                        "0x65f5": "ImapInternalDate",
                        "0x6751": "NextArticleId",
                        "0x6752": "ImapLastArticleId",
                        "0x0e2f": "ImapId",
                        "0x0e32": "OriginalSourceServerVersion",
                        "0x5806": "DeliverAsRead",
                        "0x67fe": "ReadCn",
                        "0x6808": "EventMask",
                        "0x676a": "EventMailboxGuid",
                        "0x6815": "DocumentId",
                        "0x6702": "BeingDeleted",
                        "0x678d": "FolderCDN",
                        "0x67f6": "ModifiedCount",
                        "0x67f7": "DeleteState",
                        "0x66fe": "AdminDisplayName",
                        "0x66a9": "LastAccessTime",
                        "0x6830": "LastUserAccessTime",
                        "0x6831": "LastUserModificationTime",
                        "0x66b4": "AssocMessageSizeExtended",
                        "0x66b5": "FolderPathName",
                        "0x66b6": "OwnerCount",
                        "0x66b7": "ContactCount",
                        "0x7c01": "MessageAudioNotes",
                        "0x3ff5": "StorageQuotaLimit",
                        "0x3ff6": "ExcessStorageUsed",
                        "0x3ff7": "SvrGeneratingQuotaMsg",
                        "0x3fc2": "PrimaryMbxOverQuota",
                        "0x65c6": "SecureSubmitFlags",
                        "0x673e": "PropertyGroupInformation",
                        "0x6784": "SearchRestriction",
                        "0x67b0": "ViewRestriction",
                        "0x6788": "LCIDRestriction",
                        "0x676e": "LCID",
                        "0x67f3": "ViewAccessTime",
                        "0x689e": "CategCount",
                        "0x6819": "SoftDeletedFilter",
                        "0x681b": "ConversationsFilter",
                        "0x689c": "DVUIdLowest",
                        "0x689d": "DVUIdHighest",
                        "0x6880": "ConversationMvFrom",
                        "0x6881": "ConversationMvFromMailboxWide",
                        "0x6882": "ConversationMvTo",
                        "0x6883": "ConversationMvToMailboxWide",
                        "0x6884": "ConversationMsgDeliveryTime",
                        "0x6885": "ConversationMsgDeliveryTimeMailboxWide",
                        "0x6886": "ConversationCategories",
                        "0x6887": "ConversationCategoriesMailboxWide",
                        "0x6888": "ConversationFlagStatus",
                        "0x6889": "ConversationFlagStatusMailboxWide",
                        "0x688a": "ConversationFlagCompleteTime",
                        "0x688b": "ConversationFlagCompleteTimeMailboxWide",
                        "0x688c": "ConversationHasAttach",
                        "0x688d": "ConversationHasAttachMailboxWide",
                        "0x688e": "ConversationContentCount",
                        "0x688f": "ConversationContentCountMailboxWide",
                        "0x6893": "ConversationMessageSizeMailboxWide",
                        "0x6894": "ConversationMessageClasses",
                        "0x6895": "ConversationMessageClassesMailboxWide",
                        "0x6896": "ConversationReplyForwardState",
                        "0x6897": "ConversationReplyForwardStateMailboxWide",
                        "0x6898": "ConversationImportance",
                        "0x6899": "ConversationImportanceMailboxWide",
                        "0x689a": "ConversationMvFromUnread",
                        "0x689b": "ConversationMvFromUnreadMailboxWide",
                        "0x68a0": "ConversationMvItemIds",
                        "0x68a1": "ConversationMvItemIdsMailboxWide",
                        "0x68a2": "ConversationHasIrm",
                        "0x68a3": "ConversationHasIrmMailboxWide",
                        "0x682c": "TransportSyncSubscriptionListTimestamp",
                        "0x3690": "TransportRulesSnapshot",
                        "0x3691": "TransportRulesSnapshotId",
                        "0x7c05": "DeletedMessageSizeExtendedLastModificationTime",
                        "0x0082": "ReportOriginalSender",
                        "0x0083": "ReportDispositionToNames",
                        "0x0084": "ReportDispositionToEmailAddress",
                        "0x0085": "ReportDispositionOptions",
                        "0x0086": "RichContent",
                        "0x0100": "AdministratorEMail",
                        "0x0c24": "ParticipantSID",
                        "0x0c25": "ParticipantGuid",
                        "0x0c26": "ToGroupExpansionRecipients",
                        "0x0c27": "CcGroupExpansionRecipients",
                        "0x0c28": "BccGroupExpansionRecipients",
                        "0x0e0b": "ImmutableEntryId",
                        "0x0e2e": "MessageIsHidden",
                        "0x0e33": "OlcPopId",
                        "0x0e38": "ReplFlags",
                        "0x0e40": "SenderGuid",
                        "0x0e41": "SentRepresentingGuid",
                        "0x0e42": "OriginalSenderGuid",
                        "0x0e43": "OriginalSentRepresentingGuid",
                        "0x0e44": "ReadReceiptGuid",
                        "0x0e45": "ReportGuid",
                        "0x0e46": "OriginatorGuid",
                        "0x0e47": "ReportDestinationGuid",
                        "0x0e48": "OriginalAuthorGuid",
                        "0x0e49": "ReceivedByGuid",
                        "0x0e4a": "ReceivedRepresentingGuid",
                        "0x0e4b": "CreatorGuid",
                        "0x0e4c": "LastModifierGuid",
                        "0x0e4d": "SenderSID",
                        "0x0e4e": "SentRepresentingSID",
                        "0x0e4f": "OriginalSenderSid",
                        "0x0e50": "OriginalSentRepresentingSid",
                        "0x0e51": "ReadReceiptSid",
                        "0x0e52": "ReportSid",
                        "0x0e53": "OriginatorSid",
                        "0x0e54": "ReportDestinationSid",
                        "0x0e55": "OriginalAuthorSid",
                        "0x0e56": "ReceivedBySid",
                        "0x0e57": "ReceivedRepresentingSid",
                        "0x0e58": "CreatorSID",
                        "0x0e59": "LastModifierSid",
                        "0x0e5a": "RecipientCAI",
                        "0x0e5b": "ConversationCreatorSID",
                        "0x0e5d": "IsUserKeyDecryptPossible",
                        "0x0e5e": "MaxIndices",
                        "0x0e5f": "SourceFid",
                        "0x0e60": "PFContactsGuid",
                        "0x0e61": "UrlCompNamePostfix",
                        "0x0e62": "URLCompNameSet",
                        "0x0e64": "DeletedSubfolderCount",
                        "0x0e68": "MaxCachedViews",
                        "0x0e6b": "AdminNTSecurityDescriptorAsXML",
                        "0x0e6c": "CreatorSidAsXML",
                        "0x0e6d": "LastModifierSidAsXML",
                        "0x0e6e": "SenderSIDAsXML",
                        "0x0e6f": "SentRepresentingSidAsXML",
                        "0x0e70": "OriginalSenderSIDAsXML",
                        "0x0e71": "OriginalSentRepresentingSIDAsXML",
                        "0x0e72": "ReadReceiptSIDAsXML",
                        "0x0e73": "ReportSIDAsXML",
                        "0x0e74": "OriginatorSidAsXML",
                        "0x0e75": "ReportDestinationSIDAsXML",
                        "0x0e76": "OriginalAuthorSIDAsXML",
                        "0x0e77": "ReceivedBySIDAsXML",
                        "0x0e78": "ReceivedRepersentingSIDAsXML",
                        "0x0e7a": "MergeMidsetDeleted",
                        "0x0e7b": "ReserveRangeOfIDs",
                        "0x0e97": "AddrTo",
                        "0x0e98": "AddrCc",
                        "0x0e9f": "EntourageSentHistory",
                        "0x0ea2": "ProofInProgress",
                        "0x0ea5": "SearchAttachmentsOLK",
                        "0x0ea6": "SearchRecipEmailTo",
                        "0x0ea7": "SearchRecipEmailCc",
                        "0x0ea8": "SearchRecipEmailBcc",
                        "0x0eaa": "SFGAOFlags",
                        "0x0ece": "SearchIsPartiallyIndexed",
                        "0x0ecf": "SearchUniqueBody",
                        "0x0ed0": "SearchErrorCode",
                        "0x0ed1": "SearchReceivedTime",
                        "0x0ed2": "SearchNumberOfTopRankedResults",
                        "0x0ed3": "SearchControlFlags",
                        "0x0ed4": "SearchRankingModel",
                        "0x0ed5": "SearchMinimumNumberOfDateOrderedResults",
                        "0x0ed6": "SearchSharePointOnlineSearchableProps",
                        "0x0ed7": "SearchRelevanceRankedResults",
                        "0x0edd": "MailboxSyncState",
                        "0x0f01": "RenewTime",
                        "0x0f02": "DeliveryOrRenewTime",
                        "0x0f03": "ConversationThreadId",
                        "0x0f04": "LikeCount",
                        "0x0f05": "RichContentDeprecated",
                        "0x0f06": "PeopleCentricConversationId",
                        "0x0f07": "ReturnTime",
                        "0x0f08": "LastAttachmentsProcessedTime",
                        "0x0f0a": "LastActivityTime",
                        "0x100a": "AlternateBestBody",
                        "0x100c": "IsIntegJobCorruptions",
                        "0x100e": "IsIntegJobPriority",
                        "0x100f": "IsIntegJobTimeInServer",
                        "0x1017": "AnnotationToken",
                        "0x1030": "InternetApproved",
                        "0x1033": "InternetFollowupTo",
                        "0x1036": "InetNewsgroups",
                        "0x103d": "PostReplyFolderEntries",
                        "0x1040": "NNTPXRef",
                        "0x1084": "Relevance",
                        "0x1092": "FormatPT",
                        "0x10c0": "SMTPTempTblData",
                        "0x10c1": "SMTPTempTblData2",
                        "0x10c2": "SMTPTempTblData3",
                        "0x10f0": "IMAPCachedMsgSize",
                        "0x10f2": "DisableFullFidelity",
                        "0x10f3": "UrlCompName",
                        "0x10f5": "AttrSystem",
                        "0x1204": "PredictedActions",
                        "0x1205": "GroupingActions",
                        "0x1206": "PredictedActionsSummary",
                        "0x1207": "IsClutter",
                        "0x120b": "OriginalDeliveryFolderInfo",
                        "0x120c": "ClutterFolderEntryIdWellKnown",
                        "0x120d": "BirthdayCalendarFolderEntryIdWellKnown",
                        "0x120e": "InferencePredictedClutterReasons",
                        "0x120f": "InferencePredictedNotClutterReasons",
                        "0x1210": "BookingStaffFolderEntryId",
                        "0x1211": "BookingServicesFolderEntryId",
                        "0x1212": "InferenceClassificationInternal",
                        "0x1213": "InferenceClassification",
                        "0x1214": "SchedulesFolderEntryId",
                        "0x1215": "AllTaggedItemsFolderEntryId",
                        "0x1236": "WellKnownFolderGuid",
                        "0x1237": "RemoteFolderSyncStatus",
                        "0x1238": "BookingCustomQuestionsFolderEntryId",
                        "0x300e": "UserInformationAntispamBypassEnabled",
                        "0x300f": "UserInformationArchiveDomain",
                        "0x3017": "UserInformationBirthdate",
                        "0x3020": "UserInformationCountryOrRegion",
                        "0x3021": "UserInformationDefaultMailTip",
                        "0x3022": "UserInformationDeliverToMailboxAndForward",
                        "0x3023": "UserInformationDescription",
                        "0x3024": "UserInformationDisabledArchiveGuid",
                        "0x3025": "UserInformationDowngradeHighPriorityMessagesEnabled",
                        "0x3026": "UserInformationECPEnabled",
                        "0x3027": "UserInformationEmailAddressPolicyEnabled",
                        "0x3028": "UserInformationEwsAllowEntourage",
                        "0x3029": "UserInformationEwsAllowMacOutlook",
                        "0x302a": "UserInformationEwsAllowOutlook",
                        "0x302b": "UserInformationEwsApplicationAccessPolicy",
                        "0x302c": "UserInformationEwsEnabled",
                        "0x302d": "UserInformationEwsExceptions",
                        "0x302e": "UserInformationEwsWellKnownApplicationAccessPolicies",
                        "0x302f": "UserInformationExchangeGuid",
                        "0x3030": "UserInformationExternalOofOptions",
                        "0x3031": "UserInformationFirstName",
                        "0x3032": "UserInformationForwardingSmtpAddress",
                        "0x3033": "UserInformationGender",
                        "0x3034": "UserInformationGenericForwardingAddress",
                        "0x3035": "UserInformationGeoCoordinates",
                        "0x3036": "UserInformationHABSeniorityIndex",
                        "0x3037": "UserInformationHasActiveSyncDevicePartnership",
                        "0x3038": "UserInformationHiddenFromAddressListsEnabled",
                        "0x3039": "UserInformationHiddenFromAddressListsValue",
                        "0x303a": "UserInformationHomePhone",
                        "0x303b": "UserInformationImapEnabled",
                        "0x303c": "UserInformationImapEnableExactRFC822Size",
                        "0x303d": "UserInformationImapForceICalForCalendarRetrievalOption",
                        "0x303e": "UserInformationImapMessagesRetrievalMimeFormat",
                        "0x303f": "UserInformationImapProtocolLoggingEnabled",
                        "0x3040": "UserInformationImapSuppressReadReceipt",
                        "0x3041": "UserInformationImapUseProtocolDefaults",
                        "0x3042": "UserInformationIncludeInGarbageCollection",
                        "0x3043": "UserInformationInitials",
                        "0x3044": "UserInformationInPlaceHolds",
                        "0x3045": "UserInformationInternalOnly",
                        "0x3046": "UserInformationInternalUsageLocation",
                        "0x3047": "UserInformationInternetEncoding",
                        "0x3048": "UserInformationIsCalculatedTargetAddress",
                        "0x3049": "UserInformationIsExcludedFromServingHierarchy",
                        "0x304a": "UserInformationIsHierarchyReady",
                        "0x304b": "UserInformationIsInactiveMailbox",
                        "0x304c": "UserInformationIsSoftDeletedByDisable",
                        "0x304d": "UserInformationIsSoftDeletedByRemove",
                        "0x304e": "UserInformationIssueWarningQuota",
                        "0x304f": "UserInformationJournalArchiveAddress",
                        "0x3051": "UserInformationLastExchangeChangedTime",
                        "0x3052": "UserInformationLastName",
                        "0x3053": "UserInformationLastAliasSyncSubmittedTime",
                        "0x3054": "UserInformationLEOEnabled",
                        "0x3055": "UserInformationLocaleID",
                        "0x3056": "UserInformationLongitude",
                        "0x3057": "UserInformationMacAttachmentFormat",
                        "0x3058": "UserInformationMailboxContainerGuid",
                        "0x3059": "UserInformationMailboxMoveBatchName",
                        "0x305a": "UserInformationMailboxMoveRemoteHostName",
                        "0x305b": "UserInformationMailboxMoveStatus",
                        "0x305c": "UserInformationMailboxRelease",
                        "0x305d": "UserInformationMailTipTranslations",
                        "0x305e": "UserInformationMAPIBlockOutlookNonCachedMode",
                        "0x305f": "UserInformationMAPIBlockOutlookRpcHttp",
                        "0x3060": "UserInformationMAPIBlockOutlookVersions",
                        "0x3061": "UserInformationMailboxStatus",
                        "0x3062": "UserInformationMapiRecipient",
                        "0x3063": "UserInformationMaxBlockedSenders",
                        "0x3064": "UserInformationMaxReceiveSize",
                        "0x3065": "UserInformationMaxSafeSenders",
                        "0x3066": "UserInformationMaxSendSize",
                        "0x3067": "UserInformationMemberName",
                        "0x3068": "UserInformationMessageBodyFormat",
                        "0x3069": "UserInformationMessageFormat",
                        "0x306a": "UserInformationMessageTrackingReadStatusDisabled",
                        "0x306b": "UserInformationMobileFeaturesEnabled",
                        "0x306c": "UserInformationMobilePhone",
                        "0x306d": "UserInformationModerationFlags",
                        "0x306e": "UserInformationNotes",
                        "0x306f": "UserInformationOccupation",
                        "0x3070": "UserInformationOpenDomainRoutingDisabled",
                        "0x3071": "UserInformationOtherHomePhone",
                        "0x3072": "UserInformationOtherMobile",
                        "0x3073": "UserInformationOtherTelephone",
                        "0x3074": "UserInformationOWAEnabled",
                        "0x3075": "UserInformationOWAforDevicesEnabled",
                        "0x3076": "UserInformationPager",
                        "0x3077": "UserInformationPersistedCapabilities",
                        "0x3078": "UserInformationPhone",
                        "0x3079": "UserInformationPhoneProviderId",
                        "0x307a": "UserInformationPopEnabled",
                        "0x307b": "UserInformationPopEnableExactRFC822Size",
                        "0x307c": "UserInformationPopForceICalForCalendarRetrievalOption",
                        "0x307d": "UserInformationPopMessagesRetrievalMimeFormat",
                        "0x307e": "UserInformationPopProtocolLoggingEnabled",
                        "0x307f": "UserInformationPopSuppressReadReceipt",
                        "0x3080": "UserInformationPopUseProtocolDefaults",
                        "0x3081": "UserInformationPostalCode",
                        "0x3082": "UserInformationPostOfficeBox",
                        "0x3083": "UserInformationPreviousExchangeGuid",
                        "0x3084": "UserInformationPreviousRecipientTypeDetails",
                        "0x3085": "UserInformationProhibitSendQuota",
                        "0x3086": "UserInformationProhibitSendReceiveQuota",
                        "0x3087": "UserInformationQueryBaseDNRestrictionEnabled",
                        "0x3088": "UserInformationRecipientDisplayType",
                        "0x3089": "UserInformationRecipientLimits",
                        "0x308a": "UserInformationRecipientSoftDeletedStatus",
                        "0x308b": "UserInformationRecoverableItemsQuota",
                        "0x308c": "UserInformationRecoverableItemsWarningQuota",
                        "0x308d": "UserInformationRegion",
                        "0x308e": "UserInformationRemotePowerShellEnabled",
                        "0x308f": "UserInformationRemoteRecipientType",
                        "0x3090": "UserInformationRequireAllSendersAreAuthenticated",
                        "0x3091": "UserInformationResetPasswordOnNextLogon",
                        "0x3092": "UserInformationRetainDeletedItemsFor",
                        "0x3093": "UserInformationRetainDeletedItemsUntilBackup",
                        "0x3094": "UserInformationRulesQuota",
                        "0x3095": "UserInformationShouldUseDefaultRetentionPolicy",
                        "0x3096": "UserInformationSimpleDisplayName",
                        "0x3097": "UserInformationSingleItemRecoveryEnabled",
                        "0x3098": "UserInformationStateOrProvince",
                        "0x3099": "UserInformationStreetAddress",
                        "0x309a": "UserInformationSubscriberAccessEnabled",
                        "0x309b": "UserInformationTextEncodedORAddress",
                        "0x309c": "UserInformationTextMessagingState",
                        "0x309d": "UserInformationTimezone",
                        "0x309e": "UserInformationUCSImListMigrationCompleted",
                        "0x309f": "UserInformationUpgradeDetails",
                        "0x30a0": "UserInformationUpgradeMessage",
                        "0x30a1": "UserInformationUpgradeRequest",
                        "0x30a2": "UserInformationUpgradeStage",
                        "0x30a3": "UserInformationUpgradeStageTimeStamp",
                        "0x30a4": "UserInformationUpgradeStatus",
                        "0x30a5": "UserInformationUsageLocation",
                        "0x30a6": "UserInformationUseMapiRichTextFormat",
                        "0x30a7": "UserInformationUsePreferMessageFormat",
                        "0x30a8": "UserInformationUseUCCAuditConfig",
                        "0x30a9": "UserInformationWebPage",
                        "0x30aa": "UserInformationWhenMailboxCreated",
                        "0x30ab": "UserInformationWhenSoftDeleted",
                        "0x30ac": "UserInformationBirthdayPrecision",
                        "0x30ad": "UserInformationNameVersion",
                        "0x30ae": "UserInformationOptInTime",
                        "0x30af": "UserInformationIsMigratedConsumerMailbox",
                        "0x30b0": "UserInformationMigrationDryRun",
                        "0x30b1": "UserInformationIsPremiumConsumerMailbox",
                        "0x30b2": "UserInformationAlternateSupportEmailAddresses",
                        "0x30b3": "UserInformationEmailAddresses",
                        "0x30b4": "UserInformationHasSnackyAppData",
                        "0x30b5": "UserInformationMailboxMoveTargetMDB",
                        "0x30b6": "UserInformationMailboxMoveSourceMDB",
                        "0x30b7": "UserInformationMailboxMoveFlags",
                        "0x30b8": "UserInformationHydraLastSyncTimestamp",
                        "0x30b9": "UserInformationHydraSyncStartIdentity",
                        "0x30ba": "UserInformationHydraSyncStartTimestamp",
                        "0x30bb": "UserInformationStatus",
                        "0x30bc": "UserInformationDeletedOn",
                        "0x30bd": "UserInformationMigrationInterruptionTest",
                        "0x30be": "UserInformationLocatorSource",
                        "0x30bf": "UserInformationMAPIEnabled",
                        "0x30c0": "UserInformationOlcDatFlags",
                        "0x30c1": "UserInformationOlcDat2Flags",
                        "0x30c2": "UserInformationDefaultFromAddress",
                        "0x30c3": "UserInformationNotManagedEmailAddresses",
                        "0x30c4": "UserInformationLatitude",
                        "0x30c5": "UserInformationConnectedAccounts",
                        "0x30c6": "UserInformationAccountTrustLevel",
                        "0x30c7": "UserInformationBlockReason",
                        "0x30c8": "UserInformationHijackDetection",
                        "0x30c9": "UserInformationHipChallengeApplicable",
                        "0x30ca": "UserInformationIsBlocked",
                        "0x30cb": "UserInformationIsSwitchUser",
                        "0x30cc": "UserInformationIsToolsAccount",
                        "0x30cd": "UserInformationLastBlockTime",
                        "0x30ce": "UserInformationMaxDailyMessages",
                        "0x30cf": "UserInformationReportToExternalSender",
                        "0x30d0": "UserInformationWhenOlcMailboxCreated",
                        "0x30d1": "UserInformationMailboxProvisioningConstraint",
                        "0x30d2": "UserInformationMailboxProvisioningPreferences",
                        "0x30d3": "UserInformationCID",
                        "0x30d4": "UserInformationSharingAnonymousIdentities",
                        "0x30d5": "UserInformationExchangeSecurityDescriptor",
                        "0x30d6": "UserInformationMapiHttpEnabled",
                        "0x30d7": "UserInformationMAPIBlockOutlookExternalConnectivity",
                        "0x30d8": "UserInformationUniversalOutlookEnabled",
                        "0x30d9": "UserInformationPopMessageDeleteEnabled",
                        "0x30da": "UserInformationPrimaryMailboxSource",
                        "0x30db": "UserInformationLocatorCacheHint",
                        "0x30dc": "UserInformationNetID",
                        "0x30dd": "UserInformationIsProsumerConsumerMailbox",
                        "0x30de": "UserInformationProsumerEmailAddresses",
                        "0x30df": "UserInformationProsumerMSAVerifiedEmailAddresses",
                        "0x30e0": "UserInformationShardOwnerExchangeObjectId",
                        "0x30e1": "UserInformationShardOwnerTenantPartitionHint",
                        "0x30e2": "UserInformationShardProvisionedByProtocolType",
                        "0x30e3": "UserInformationIsShadowMailboxProvisioningComplete",
                        "0x30e4": "UserInformationShadowRemoteEmailAddress",
                        "0x30e5": "UserInformationShadowScope",
                        "0x30e6": "UserInformationShadowUserName",
                        "0x30e7": "UserInformationShadowProvider",
                        "0x30e8": "UserInformationIsShadowMailbox",
                        "0x30e9": "UserInformationPersistedMservNameVersion",
                        "0x30ea": "UserInformationLastPersistedMservNameVersionUpdateTime",
                        "0x30eb": "UserInformationPremiumAccountOffers",
                        "0x30ec": "UserInformationLegacyCustomDomainAddresses",
                        "0x30ed": "UserInformationActiveSyncSuppressReadReceipt",
                        "0x30ee": "UserInformationAcceptMessagesOnlyFrom",
                        "0x30ef": "UserInformationAcceptMessagesOnlyFromBL",
                        "0x30f0": "UserInformationAcceptMessagesOnlyFromDLMembers",
                        "0x30f1": "UserInformationAcceptMessagesOnlyFromDLMembersBL",
                        "0x30f2": "UserInformationActiveSyncMailboxPolicy",
                        "0x30f3": "UserInformationActiveSyncMailboxPolicyIsDefaulted",
                        "0x30f4": "UserInformationAddressBookFlags",
                        "0x30f5": "UserInformationAddressBookPolicy",
                        "0x30f6": "UserInformationAddressListMembership",
                        "0x30f7": "UserInformationAdministrativeUnits",
                        "0x30f8": "UserInformationAggregatedMailboxGuidsRaw",
                        "0x30f9": "UserInformationAlias",
                        "0x30fa": "UserInformationAllowAddGuests",
                        "0x30fb": "UserInformationAllowedAttributesEffective",
                        "0x30fc": "UserInformationAllowUMCallsFromNonUsers",
                        "0x30fd": "UserInformationAltSecurityIdentities",
                        "0x30fe": "UserInformationApprovalApplications",
                        "0x30ff": "UserInformationArbitrationMailbox",
                        "0x3100": "UserInformationArchiveDatabaseRaw",
                        "0x3101": "UserInformationAttributeMetadata",
                        "0x3102": "UserInformationAuditAdminFlags",
                        "0x3103": "UserInformationAuditBypassEnabled",
                        "0x3104": "UserInformationAuditDelegateAdminFlags",
                        "0x3105": "UserInformationAuditDelegateFlags",
                        "0x3106": "UserInformationAuditEnabled",
                        "0x3107": "UserInformationAuditLastAdminAccess",
                        "0x3108": "UserInformationAuditLastDelegateAccess",
                        "0x3109": "UserInformationAuditLastExternalAccess",
                        "0x310a": "UserInformationAuditLogAgeLimit",
                        "0x310b": "UserInformationAuditOwnerFlags",
                        "0x310c": "UserInformationAuditStorageState",
                        "0x310d": "UserInformationAuxMailboxParentObjectId",
                        "0x310e": "UserInformationAuxMailboxParentObjectIdBL",
                        "0x310f": "UserInformationAuthenticationPolicy",
                        "0x3110": "UserInformationBlockedSendersHash",
                        "0x3111": "UserInformationBypassModerationFrom",
                        "0x3112": "UserInformationBypassModerationFromBL",
                        "0x3113": "UserInformationBypassModerationFromDLMembers",
                        "0x3114": "UserInformationBypassModerationFromDLMembersBL",
                        "0x3115": "UserInformationCallAnsweringAudioCodecLegacy",
                        "0x3116": "UserInformationCallAnsweringAudioCodec2",
                        "0x3117": "UserInformationCatchAllRecipientBL",
                        "0x3118": "UserInformationCertificate",
                        "0x3119": "UserInformationClassification",
                        "0x311a": "UserInformationCo",
                        "0x311b": "UserInformationCoManagedBy",
                        "0x311c": "UserInformationCoManagedObjectsBL",
                        "0x311d": "UserInformationCompany",
                        "0x311e": "UserInformationConfigurationUnit",
                        "0x311f": "UserInformationConfigurationXMLRaw",
                        "0x3120": "UserInformationCorrelationIdRaw",
                        "0x3121": "UserInformationCustomAttribute1",
                        "0x3122": "UserInformationCustomAttribute10",
                        "0x3123": "UserInformationCustomAttribute11",
                        "0x3124": "UserInformationCustomAttribute12",
                        "0x3125": "UserInformationCustomAttribute13",
                        "0x3126": "UserInformationCustomAttribute14",
                        "0x3127": "UserInformationCustomAttribute15",
                        "0x3128": "UserInformationCustomAttribute2",
                        "0x3129": "UserInformationCustomAttribute3",
                        "0x312a": "UserInformationCustomAttribute4",
                        "0x312b": "UserInformationCustomAttribute5",
                        "0x312c": "UserInformationCustomAttribute6",
                        "0x312d": "UserInformationCustomAttribute7",
                        "0x312e": "UserInformationCustomAttribute8",
                        "0x312f": "UserInformationCustomAttribute9",
                        "0x3130": "UserInformationDatabase",
                        "0x3131": "UserInformationDataEncryptionPolicy",
                        "0x3132": "UserInformationDefaultPublicFolderMailbox",
                        "0x3133": "UserInformationDefaultPublicFolderMailboxSmtpAddress",
                        "0x3134": "UserInformationDelegateListBL",
                        "0x3135": "UserInformationDelegateListLink",
                        "0x3136": "UserInformationDeletedItemFlags",
                        "0x3137": "UserInformationDeliveryMechanism",
                        "0x3138": "UserInformationDepartment",
                        "0x3139": "UserInformationDirectReports",
                        "0x313a": "UserInformationDirSyncAuthorityMetadata",
                        "0x313b": "UserInformationDirSyncId",
                        "0x313c": "UserInformationDisabledArchiveDatabase",
                        "0x313d": "UserInformationDLSupervisionList",
                        "0x313e": "UserInformationElcExpirationSuspensionEndDate",
                        "0x313f": "UserInformationElcExpirationSuspensionStartDate",
                        "0x3140": "UserInformationElcMailboxFlags",
                        "0x3141": "UserInformationElcPolicyTemplate",
                        "0x3142": "UserInformationEntryId",
                        "0x3143": "UserInformationExchangeObjectIdRaw",
                        "0x3144": "UserInformationExchangeSecurityDescriptorRaw",
                        "0x3145": "UserInformationExchangeVersion",
                        "0x3146": "UserInformationExchangeUserAccountControl",
                        "0x3147": "UserInformationExpansionServer",
                        "0x3148": "UserInformationExtensionCustomAttribute1",
                        "0x3149": "UserInformationExtensionCustomAttribute2",
                        "0x314a": "UserInformationExtensionCustomAttribute3",
                        "0x314b": "UserInformationExtensionCustomAttribute4",
                        "0x314c": "UserInformationExtensionCustomAttribute5",
                        "0x314d": "UserInformationExternalDirectoryObjectId",
                        "0x314e": "UserInformationExternalSyncState",
                        "0x314f": "UserInformationFax",
                        "0x3150": "UserInformationFblEnabled",
                        "0x3151": "UserInformationForwardingAddress",
                        "0x3152": "UserInformationForwardingAddressBL",
                        "0x3153": "UserInformationForeignGroupSid",
                        "0x3154": "UserInformationGeneratedOfflineAddressBooks",
                        "0x3155": "UserInformationGroupPersonification",
                        "0x3156": "UserInformationGrantSendOnBehalfTo",
                        "0x3157": "UserInformationGrantSendOnBehalfToBL",
                        "0x3158": "UserInformationGroupSubtypeName",
                        "0x3159": "UserInformationGroupType",
                        "0x315a": "UserInformationGroupExternalMemberCount",
                        "0x315b": "UserInformationGroupMemberCount",
                        "0x315c": "UserInformationGuestHint",
                        "0x315d": "UserInformationHABShowInDepartments",
                        "0x315e": "UserInformationHeuristics",
                        "0x315f": "UserInformationHiddenGroupMembershipEnabled",
                        "0x3160": "UserInformationHomeMTA",
                        "0x3161": "UserInformationId",
                        "0x3162": "UserInformationImmutableId",
                        "0x3163": "UserInformationInPlaceHoldsRaw",
                        "0x3164": "UserInformationIntendedMailboxPlan",
                        "0x3165": "UserInformationInternalRecipientSupervisionList",
                        "0x3166": "UserInformationIsDirSynced",
                        "0x3167": "UserInformationIsInactive",
                        "0x3168": "UserInformationIsOrganizationalGroup",
                        "0x3169": "UserInformationLdapRecipientFilter",
                        "0x316a": "UserInformationLanguagesRaw",
                        "0x316b": "UserInformationLegacyExchangeDN",
                        "0x316c": "UserInformationLinkedPartnerGroupAndOrganizationId",
                        "0x316d": "UserInformationLinkMetadata",
                        "0x316e": "UserInformationLitigationHoldDate",
                        "0x316f": "UserInformationLitigationHoldOwner",
                        "0x3170": "UserInformationLocalizationFlags",
                        "0x3171": "UserInformationMailboxDatabasesRaw",
                        "0x3172": "UserInformationMailboxGuidsRaw",
                        "0x3173": "UserInformationMailboxLocationsRaw",
                        "0x3174": "UserInformationMailboxPlan",
                        "0x3175": "UserInformationMailboxPlanIndex",
                        "0x3176": "UserInformationMailboxRegion",
                        "0x3177": "UserInformationMailboxMoveSourceArchiveMDB",
                        "0x3178": "UserInformationMailboxMoveTargetArchiveMDB",
                        "0x3179": "UserInformationMbxGuidEnabled",
                        "0x317a": "UserInformationManager",
                        "0x317b": "UserInformationMasterAccountSid",
                        "0x317c": "UserInformationMasterDirectoryObjectIdRaw",
                        "0x317d": "UserInformationMemberDepartRestriction",
                        "0x317e": "UserInformationMemberJoinRestriction",
                        "0x317f": "UserInformationMemberOfGroup",
                        "0x3180": "UserInformationMembers",
                        "0x3181": "UserInformationMessageHygieneFlags",
                        "0x3182": "UserInformationMigrationToUnifiedGroupInProgress",
                        "0x3183": "UserInformationMobileAdminExtendedSettings",
                        "0x3184": "UserInformationMobileMailboxFlags",
                        "0x3185": "UserInformationModeratedBy",
                        "0x3186": "UserInformationModerationEnabled",
                        "0x3187": "UserInformationModeratedObjectsBL",
                        "0x3188": "UserInformationMservNameVersion",
                        "0x3189": "UserInformationMservNetID",
                        "0x318a": "UserInformationNTSecurityDescriptor",
                        "0x318b": "UserInformationObjectCategory",
                        "0x318c": "UserInformationObjectClass",
                        "0x318d": "UserInformationOffice",
                        "0x318e": "UserInformationOfflineAddressBook",
                        "0x318f": "UserInformationOneOffSupervisionList",
                        "0x3190": "UserInformationOrganizationalUnitRoot",
                        "0x3191": "UserInformationOrgLeaders",
                        "0x3192": "UserInformationOriginatingServer",
                        "0x3193": "UserInformationOtherDisplayNames",
                        "0x3194": "UserInformationOtherFax",
                        "0x3195": "UserInformationOwaMailboxPolicy",
                        "0x3196": "UserInformationOwners",
                        "0x3197": "UserInformationPreviousDatabase",
                        "0x3198": "UserInformationPublicFolderContacts",
                        "0x3199": "UserInformationPuidEmailAddressEnabled",
                        "0x319a": "UserInformationPurportedSearchUI",
                        "0x319b": "UserInformationPasswordLastSetRaw",
                        "0x319c": "UserInformationPhoneticCompany",
                        "0x319d": "UserInformationPhoneticDisplayName",
                        "0x319e": "UserInformationPhoneticDepartment",
                        "0x319f": "UserInformationPhoneticFirstName",
                        "0x31a0": "UserInformationPhoneticLastName",
                        "0x31a1": "UserInformationPoliciesExcluded",
                        "0x31a2": "UserInformationPoliciesIncluded",
                        "0x31a3": "UserInformationPrimaryGroupId",
                        "0x31a4": "UserInformationProtocolSettings",
                        "0x31a5": "UserInformationProvisioningFlags",
                        "0x31a6": "UserInformationQueryBaseDN",
                        "0x31a7": "UserInformationRawCanonicalName",
                        "0x31a8": "UserInformationRawCapabilities",
                        "0x31a9": "UserInformationRawExternalEmailAddress",
                        "0x31aa": "UserInformationRawManagedBy",
                        "0x31ab": "UserInformationRawName",
                        "0x31ac": "UserInformationRawDisplayName",
                        "0x31ad": "UserInformationRawOnPremisesObjectId",
                        "0x31ae": "UserInformationRecipientContainer",
                        "0x31af": "UserInformationRecipientFilter",
                        "0x31b0": "UserInformationRecipientFilterMetadata",
                        "0x31b1": "UserInformationRecipientTypeDetailsValue",
                        "0x31b2": "UserInformationReplicationSignature",
                        "0x31b3": "UserInformationRejectMessagesFrom",
                        "0x31b4": "UserInformationRejectMessagesFromBL",
                        "0x31b5": "UserInformationRejectMessagesFromDLMembers",
                        "0x31b6": "UserInformationRejectMessagesFromDLMembersBL",
                        "0x31b7": "UserInformationReleaseTrack",
                        "0x31b8": "UserInformationRemoteAccountPolicy",
                        "0x31b9": "UserInformationReportToManagerEnabled",
                        "0x31ba": "UserInformationReportToOriginatorEnabled",
                        "0x31bb": "UserInformationResourceCapacity",
                        "0x31bc": "UserInformationResourceMetaData",
                        "0x31bd": "UserInformationResourcePropertiesDisplay",
                        "0x31be": "UserInformationResourceSearchProperties",
                        "0x31bf": "UserInformationRetentionComment",
                        "0x31c0": "UserInformationRetentionUrl",
                        "0x31c1": "UserInformationRMSComputerAccounts",
                        "0x31c2": "UserInformationRoleAssignmentPolicy",
                        "0x31c3": "UserInformationRoleGroupTypeId",
                        "0x31c4": "UserInformationRTCSIPPrimaryUserAddress",
                        "0x31c5": "UserInformationRtcSipLine",
                        "0x31c6": "UserInformationSafeRecipientsHash",
                        "0x31c7": "UserInformationSafeSendersHash",
                        "0x31c8": "UserInformationSamAccountName",
                        "0x31c9": "UserInformationSatchmoClusterIp",
                        "0x31ca": "UserInformationSatchmoDGroup",
                        "0x31cb": "UserInformationSCLDeleteThresholdInt",
                        "0x31cc": "UserInformationSCLJunkThresholdInt",
                        "0x31cd": "UserInformationSCLQuarantineThresholdInt",
                        "0x31ce": "UserInformationSCLRejectThresholdInt",
                        "0x31cf": "UserInformationSecurityProtocol",
                        "0x31d0": "UserInformationSendOofMessageToOriginatorEnabled",
                        "0x31d1": "UserInformationServerLegacyDN",
                        "0x31d2": "UserInformationSharePointLinkedBy",
                        "0x31d3": "UserInformationSharePointResources",
                        "0x31d4": "UserInformationSharePointSiteInfo",
                        "0x31d5": "UserInformationSharePointUrl",
                        "0x31d6": "UserInformationSharingPartnerIdentitiesRaw",
                        "0x31d7": "UserInformationSharingPolicy",
                        "0x31d8": "UserInformationSid",
                        "0x31d9": "UserInformationSidHistory",
                        "0x31da": "UserInformationSiloName",
                        "0x31db": "UserInformationSkypeId",
                        "0x31dc": "UserInformationSMimeCertificate",
                        "0x31dd": "UserInformationSourceAnchor",
                        "0x31de": "UserInformationStsRefreshTokensValidFrom",
                        "0x31df": "UserInformationSystemMailboxRetainDeletedItemsFor",
                        "0x31e0": "UserInformationSystemMailboxRulesQuota",
                        "0x31e1": "UserInformationTeamMailboxExpiration",
                        "0x31e2": "UserInformationTeamMailboxShowInClientList",
                        "0x31e3": "UserInformationTelephoneAssistant",
                        "0x31e4": "UserInformationThrottlingPolicy",
                        "0x31e5": "UserInformationThumbnailPhoto",
                        "0x31e6": "UserInformationTitle",
                        "0x31e7": "UserInformationTokenGroupsGlobalAndUniversal",
                        "0x31e8": "UserInformationTransportSettingFlags",
                        "0x31e9": "UserInformationUMAddresses",
                        "0x31ea": "UserInformationUMCallingLineIds",
                        "0x31eb": "UserInformationUMDtmfMap",
                        "0x31ec": "UserInformationUMEnabledFlags",
                        "0x31ed": "UserInformationUMEnabledFlags2",
                        "0x31ee": "UserInformationUMMailboxPolicy",
                        "0x31ef": "UserInformationUMPinChecksum",
                        "0x31f0": "UserInformationUMRecipientDialPlanId",
                        "0x31f1": "UserInformationUMServerWritableFlags",
                        "0x31f2": "UserInformationUMSpokenName",
                        "0x31f3": "UserInformationUnifiedGroupEventSubscriptionBL",
                        "0x31f4": "UserInformationUnifiedGroupEventSubscriptionLink",
                        "0x31f5": "UserInformationUnifiedGroupFileNotificationsSettings",
                        "0x31f6": "UserInformationUnifiedGroupMembersBL",
                        "0x31f7": "UserInformationUnifiedGroupMembersLink",
                        "0x31f8": "UserInformationUnifiedGroupProvisioningOption",
                        "0x31f9": "UserInformationUnifiedGroupSecurityFlags",
                        "0x31fa": "UserInformationUnifiedGroupSKU",
                        "0x31fb": "UserInformationUnifiedMailboxAccount",
                        "0x31fc": "UserInformationUserAccountControl",
                        "0x31fd": "UserInformationUserPrincipalNameRaw",
                        "0x31fe": "UserInformationUseDatabaseQuotaDefaults",
                        "0x31ff": "UserInformationUsnChanged",
                        "0x3200": "UserInformationUsnCreated",
                        "0x3201": "UserInformationUserState",
                        "0x3202": "UserInformationVoiceMailSettings",
                        "0x3203": "UserInformationWhenChangedRaw",
                        "0x3204": "UserInformationWhenCreatedRaw",
                        "0x3205": "UserInformationWindowsEmailAddress",
                        "0x3206": "UserInformationWindowsLiveID",
                        "0x3207": "UserInformationYammerGroupAddress",
                        "0x3208": "UserInformationOperatorNumber",
                        "0x3209": "UserInformationWhenReadUTC",
                        "0x320a": "UserInformationPreviousRecipientTypeDetailsHigh",
                        "0x320b": "UserInformationRemoteRecipientTypeHigh",
                        "0x320c": "UserInformationRecipientTypeDetailsValueHigh",
                        "0x320d": "UserInformationFamilyMembersUpdateInProgressStartTime",
                        "0x320e": "UserInformationIsFamilyMailbox",
                        "0x320f": "UserInformationMailboxRegionLastUpdateTime",
                        "0x3210": "UserInformationSubscribeExistingGroupMembersStatus",
                        "0x3211": "UserInformationGroupMembers",
                        "0x3212": "UserInformationRecipientDisplayTypeRaw",
                        "0x3213": "UserInformationUITEntryVersion",
                        "0x3214": "UserInformationLastRefreshedFrom",
                        "0x3215": "UserInformationIsGroupMailBox",
                        "0x3216": "UserInformationMailboxFolderSet",
                        "0x3217": "UserInformationWasInactiveMailbox",
                        "0x3218": "UserInformationInactiveMailboxRetireTime",
                        "0x3219": "UserInformationOrphanSoftDeleteTrackingTime",
                        "0x321a": "UserInformationSubscriptions",
                        "0x321b": "UserInformationOtherMail",
                        "0x321c": "UserInformationIsCIDAddedToMserv",
                        "0x321d": "UserInformationMailboxWorkloads",
                        "0x321e": "UserInformationCacheLastAccessTime",
                        "0x3233": "UserInformationPublicFolderClientAccess",
                        "0x330b": "BigFunnelLargePOITableTotalPages",
                        "0x330c": "BigFunnelLargePOITableAvailablePages",
                        "0x330d": "BigFunnelPOISize",
                        "0x330e": "BigFunnelMessageCount",
                        "0x330f": "FastIsEnabled",
                        "0x3310": "NeedsToMove",
                        "0x3311": "MCDBMessageTableTotalPages",
                        "0x3312": "MCDBMessageTableAvailablePages",
                        "0x3313": "MCDBOtherTablesTotalPages",
                        "0x3314": "MCDBOtherTablesAvailablePages",
                        "0x3315": "MCDBBigFunnelFilterTableTotalPages",
                        "0x3316": "MCDBBigFunnelFilterTableAvailablePages",
                        "0x3317": "MCDBBigFunnelLargePOITableTotalPages",
                        "0x3318": "MCDBBigFunnelLargePOITableAvailablePages",
                        "0x3319": "MCDBSize",
                        "0x3320": "MCDBAvailableSpace",
                        "0x3321": "MCDBBigFunnelPostingListTableTotalPages",
                        "0x3322": "MCDBBigFunnelPostingListTableAvailablePages",
                        "0x3323": "MCDBMessageTablePercentReplicated",
                        "0x3324": "MCDBBigFunnelFilterTablePercentReplicated",
                        "0x3325": "MCDBBigFunnelLargePOITablePercentReplicated",
                        "0x3326": "MCDBBigFunnelPostingListTablePercentReplicated",
                        "0x3327": "BigFunnelMailboxCreationVersion",
                        "0x3328": "BigFunnelAttributeVectorCommonVersion",
                        "0x3329": "BigFunnelAttributeVectorSharePointVersion",
                        "0x3330": "BigFunnelIndexedSize",
                        "0x3331": "BigFunnelPartiallyIndexedSize",
                        "0x3332": "BigFunnelNotIndexedSize",
                        "0x3333": "BigFunnelCorruptedSize",
                        "0x3334": "BigFunnelStaleSize",
                        "0x3335": "BigFunnelShouldNotBeIndexedSize",
                        "0x3336": "BigFunnelIndexedCount",
                        "0x3337": "BigFunnelPartiallyIndexedCount",
                        "0x3338": "BigFunnelNotIndexedCount",
                        "0x3339": "BigFunnelCorruptedCount",
                        "0x333a": "BigFunnelStaleCount",
                        "0x333b": "BigFunnelShouldNotBeIndexedCount",
                        "0x333c": "BigFunnelL1Rank",
                        "0x333d": "BigFunnelResultSets",
                        "0x333e": "BigFunnelMaintainRefiners",
                        "0x333f": "BigFunnelPostingListTableBuckets",
                        "0x3340": "BigFunnelPostingListTargetTableBuckets",
                        "0x3341": "BigFunnelL1FeatureNames",
                        "0x3342": "BigFunnelL1FeatureValues",
                        "0x3343": "MCDBLogonScenarioTotalPages",
                        "0x3344": "MCDBLogonScenarioAvailablePages",
                        "0x3345": "BigFunnelMasterIndexVersion",
                        "0x33f0": "ControlDataForRecordReviewNotificationTBA",
                        "0x33fe": "ControlDataForBigFunnelQueryParityAssistant",
                        "0x33ff": "BigFunnelQueryParityAssistantVersion",
                        "0x3401": "MessageTableTotalPages",
                        "0x3402": "MessageTableAvailablePages",
                        "0x3403": "OtherTablesTotalPages",
                        "0x3404": "OtherTablesAvailablePages",
                        "0x3405": "AttachmentTableTotalPages",
                        "0x3406": "AttachmentTableAvailablePages",
                        "0x3407": "MailboxTypeVersion",
                        "0x3408": "MailboxPartitionMailboxGuids",
                        "0x3409": "BigFunnelFilterTableTotalPages",
                        "0x340a": "BigFunnelFilterTableAvailablePages",
                        "0x340b": "BigFunnelPostingListTableTotalPages",
                        "0x340c": "BigFunnelPostingListTableAvailablePages",
                        "0x3417": "ProviderDisplayIcon",
                        "0x3418": "ProviderDisplayName",
                        "0x3432": "ControlDataForDirectoryProcessorAssistant",
                        "0x3433": "NeedsDirectoryProcessor",
                        "0x3434": "RetentionQueryIds",
                        "0x3435": "RetentionQueryInfo",
                        "0x3436": "MailboxLastProcessedTimestamp",
                        "0x3437": "ControlDataForPublicFolderAssistant",
                        "0x3438": "ControlDataForInferenceTrainingAssistant",
                        "0x3439": "InferenceEnabled",
                        "0x343b": "ContactLinking",
                        "0x343c": "ControlDataForOABGeneratorAssistant",
                        "0x343d": "ContactSaveVersion",
                        "0x3440": "PushNotificationSubscriptionType",
                        "0x3442": "ControlDataForInferenceDataCollectionAssistant",
                        "0x3443": "InferenceDataCollectionProcessingState",
                        "0x3444": "ControlDataForPeopleRelevanceAssistant",
                        "0x3445": "SiteMailboxInternalState",
                        "0x3446": "ControlDataForSiteMailboxAssistant",
                        "0x3447": "InferenceTrainingLastContentCount",
                        "0x3448": "InferenceTrainingLastAttemptTimestamp",
                        "0x3449": "InferenceTrainingLastSuccessTimestamp",
                        "0x344a": "InferenceUserCapabilityFlags",
                        "0x344b": "ControlDataForMailboxAssociationReplicationAssistant",
                        "0x344c": "MailboxAssociationNextReplicationTime",
                        "0x344d": "MailboxAssociationProcessingFlags",
                        "0x344e": "ControlDataForSharePointSignalStoreAssistant",
                        "0x344f": "ControlDataForPeopleCentricTriageAssistant",
                        "0x3450": "NotificationBrokerSubscriptions",
                        "0x3452": "ElcLastRunTotalProcessingTime",
                        "0x3453": "ElcLastRunSubAssistantProcessingTime",
                        "0x3454": "ElcLastRunUpdatedFolderCount",
                        "0x3455": "ElcLastRunTaggedFolderCount",
                        "0x3456": "ElcLastRunUpdatedItemCount",
                        "0x3457": "ElcLastRunTaggedWithArchiveItemCount",
                        "0x3458": "ElcLastRunTaggedWithExpiryItemCount",
                        "0x3459": "ElcLastRunDeletedFromRootItemCount",
                        "0x345a": "ElcLastRunDeletedFromDumpsterItemCount",
                        "0x345b": "ElcLastRunArchivedFromRootItemCount",
                        "0x345c": "ElcLastRunArchivedFromDumpsterItemCount",
                        "0x345d": "ScheduledISIntegLastFinished",
                        "0x345f": "ELCLastSuccessTimestamp",
                        "0x3460": "EventEmailReminderTimer",
                        "0x3463": "ControlDataForGroupMailboxAssistant",
                        "0x3464": "ItemsPendingUpgrade",
                        "0x3465": "ConsumerSharingCalendarSubscriptionCount",
                        "0x3466": "GroupMailboxGeneratedPhotoVersion",
                        "0x3467": "GroupMailboxGeneratedPhotoSignature",
                        "0x3468": "AadGroupPublishedVersion",
                        "0x3469": "ElcFaiSaveStatus",
                        "0x346a": "ElcFaiDeleteStatus",
                        "0x346b": "ControlDataForCleanupActionsAssistant",
                        "0x346c": "HolidayCalendarVersionInfo",
                        "0x346d": "HolidayCalendarSubscriptionCount",
                        "0x346e": "HolidayCalendarHierarchyVersion",
                        "0x346f": "CalendarVersion",
                        "0x3470": "ControlDataForDefaultViewIndexAssistant",
                        "0x3471": "DefaultViewAssistantLastIndexTime",
                        "0x3472": "ControlDataForAuditTimeBasedAssitantAssistant",
                        "0x3473": "SystemCategoriesViewLastIndexTime",
                        "0x3474": "ControlDataForComplianceJobAssistant",
                        "0x3475": "ControlDataForGoLocalAssistant",
                        "0x3476": "ControlDataForUserGroupsRelevanceAssistant",
                        "0x3477": "GroupMailboxLastUsageCollectionTime",
                        "0x3478": "EventPushReminderTimer",
                        "0x3479": "PushReminderSubscriptionType",
                        "0x347a": "ControlDataForHashtagsRelevanceAssistant",
                        "0x347b": "ControlDataForSearchFeatureExtractionAssistant",
                        "0x347c": "EventMeetingConflictTimer",
                        "0x347d": "MentionsViewLastIndexTime",
                        "0x347e": "O365SuiteNotificationType",
                        "0x347f": "ControlDataForResourceUsageLoggingTimeBasedAssistant",
                        "0x3480": "ControlDataForConferenceRoomUsageAssistant",
                        "0x3481": "ControlDataForRetrospectiveFeaturizationAssistant",
                        "0x3482": "ConferenceRoomUsageUpdateNeeded",
                        "0x3483": "ControlDataForGriffinTimeBasedAssistant",
                        "0x3484": "FeaturizerExperimentId",
                        "0x3485": "ControlDataForPeopleInsightsTimeBasedAssistant",
                        "0x3486": "MailboxPreferredLocation",
                        "0x3487": "ControlDataForContentSubmissionAssistant",
                        "0x3489": "ControlDataForPublicFolderHierarchySyncAssistant",
                        "0x348a": "LastActiveParentEntryId",
                        "0x348b": "WasParentDeletedItems",
                        "0x348c": "PushSyncProcessingFlags",
                        "0x348d": "ControlDataForSuggestedUserGroupAssociationsAssistant",
                        "0x348e": "ControlDataForReminderSettingsAssistant",
                        "0x348f": "ControlDataForFileExtractionTimeBasedAssistant",
                        "0x3490": "ControlDataForMailboxDataExportAssistant",
                        "0x3491": "ControlDataForXrmActivityStreamMaintenanceAssistant",
                        "0x3492": "GroupMailboxSegmentationVersion",
                        "0x3493": "ControlDataForTimeProfileTimeBasedAssistant",
                        "0x3494": "ReactorSubscriptionCreationNeeded",
                        "0x3495": "ControlDataForMailboxUsageAnalysisAssistant",
                        "0x3496": "ControlDataForContactCleanUpAssistant",
                        "0x3497": "ControlDataForGroupCalendarSubscriptionAssistant",
                        "0x3498": "ControlDataForGriffinLightweightTimeBasedAssistant",
                        "0x3499": "ControlDataForXrmAutoTaggingMaintenanceAssistant",
                        "0x349a": "ControlDataForCalculatedValueTimeBasedAssistant",
                        "0x349b": "ControlDataForSkypeContactCleanUpAssistant",
                        "0x349c": "ControlDataForShardRelevancyAssistant",
                        "0x349d": "MeetingLocationCacheVersion",
                        "0x349e": "SuperFocusedViewLastIndexTime",
                        "0x34a0": "ControlDataForCalendarFeaturizationAssistant",
                        "0x34a1": "ControlDataForBigFunnelRetryFeederTimeBasedAssistant",
                        "0x34a2": "MeetingLocationCacheVersionV3",
                        "0x34a3": "ControlDataForSharingMigrationTimeBasedAssistant",
                        "0x34a4": "ControlDataForDynamicAttachmentTimeBasedAssistant",
                        "0x34a5": "ControlDataForSharingSyncAssistant",
                        "0x34a6": "TailoredPropertiesViewLastIndexTime",
                        "0x34a7": "AtpDynamicAttachmentEnabled",
                        "0x34a8": "MailboxAssociationVersion",
                        "0x34af": "ResourceUsageAggregate",
                        "0x34b0": "ResourceUsageDataReady",
                        "0x34b1": "ResourceUsageMinDateTime",
                        "0x34b2": "ResourceUsageMaxDateTime",
                        "0x34b3": "ResourceUsageNumberOfActivities",
                        "0x34b4": "ResourceUsageNumberOfCallsSlow",
                        "0x34b5": "ResourceUsageTotalCalls",
                        "0x34b6": "ResourceUsageTotalChunks",
                        "0x34b7": "ResourceUsageTotalCpuTimeKernel",
                        "0x34b8": "ResourceUsageTotalCpuTimeUser",
                        "0x34b9": "ResourceUsageTotalDatabaseReadWaitTime",
                        "0x34ba": "ResourceUsageTotalDatabaseTime",
                        "0x34bb": "ResourceUsageTotalLogBytes",
                        "0x34bc": "ResourceUsageTotalPagesDirtied",
                        "0x34bd": "ResourceUsageTotalPagesPreread",
                        "0x34be": "ResourceUsageTotalPagesRead",
                        "0x34bf": "ResourceUsageTotalPagesRedirtied",
                        "0x34c0": "ResourceUsageTotalTime",
                        "0x34c1": "ControlDataForPicwAssistant",
                        "0x34c2": "ResourceUsageClientTypeBitmap",
                        "0x34c3": "ConnectorConfigurationCount",
                        "0x34c4": "ResourceUsageRollingAvgRopAggregate",
                        "0x34c5": "ResourceUsageRollingAvgRop",
                        "0x34c6": "ResourceUsageRollingClientTypes",
                        "0x34c7": "ControlDataForComposeGroupSuggestionTimeBasedAssistant",
                        "0x34c8": "HasSubstrateData",
                        "0x34c9": "ControlDataForAddressListIndexAssistant",
                        "0x34ca": "ControlDataForActivitySharingTimeBasedAssistant",
                        "0x34cb": "ControlDataForShardRelevancyMultiStepAssistant",
                        "0x34cc": "ControlDataForXrmProvisioningTimeBasedAssistant",
                        "0x34cd": "ControlDataForSmbTenantProvisioningAssistant",
                        "0x34ce": "ControlDataForSupervisoryReviewTimeBasedAssistant",
                        "0x34cf": "ControlDataForMailboxQuotaAssistant",
                        "0x34d0": "IsQuotaSetByMailboxQuotaAssistant",
                        "0x34d1": "SuperReactClientViewLastIndexTime",
                        "0x34d2": "FileFolderFlags",
                        "0x3500": "ControlDataForBookingsTimeBasedAssistant",
                        "0x35d8": "RootEntryId",
                        "0x35e1": "IpmInboxEntryId",
                        "0x35e8": "SpoolerQueueEntryId",
                        "0x35e9": "ProtectedMailboxKey",
                        "0x35ea": "SyncRootFolderEntryId",
                        "0x35eb": "UMVoicemailFolderEntryId",
                        "0x35ed": "EHAMigrationFolderEntryId",
                        "0x35f6": "DeletionsFolderEntryId",
                        "0x35f7": "PurgesFolderEntryId",
                        "0x35f8": "DiscoveryHoldsFolderEntryId",
                        "0x35f9": "VersionsFolderEntryId",
                        "0x35fa": "ControlDataForBigFunnelMetricsCollectionAssistant",
                        "0x35fb": "BigFunnelMetricsCollectionAssistantVersion",
                        "0x35fc": "PublicFolderDiscoveryHoldsEntryId",
                        "0x35fd": "SystemFolderEntryId",
                        "0x35ff": "ArchiveFolderEntryId",
                        "0x361c": "PackedNamedProps",
                        "0x3645": "PartOfContentIndexing",
                        "0x3647": "SearchFolderAgeOutTimeout",
                        "0x3648": "SearchFolderPopulationResult",
                        "0x3649": "SearchFolderPopulationDiagnostics",
                        "0x364a": "FolderDatabaseVersion",
                        "0x364b": "SystemMessageCount",
                        "0x364c": "SystemMessageSize",
                        "0x364d": "SystemMessageSizeWarningQuota",
                        "0x364e": "SystemMessageSizeShutoffQuota",
                        "0x364f": "TotalPages",
                        "0x3651": "ClusterMessages",
                        "0x3652": "MessageTenantPartitionHintForValidation",
                        "0x3653": "BigFunnelPOIUncompressed",
                        "0x3655": "BigFunnelPOIIsUpToDate",
                        "0x3656": "AggressiveOportunisticPromotionForMessages",
                        "0x3657": "BigFunnelPartialPOIUncompressed",
                        "0x3658": "MessageMailboxGuidForValidation",
                        "0x3659": "DatabaseSchemaVersion",
                        "0x365a": "BigFunnelPoiNotNeededReason",
                        "0x365b": "PerUserTrackingBasedOnImmutableId",
                        "0x365c": "SetSearchCriteriaFlags",
                        "0x365d": "LargeOnPageThreshold",
                        "0x3661": "SecondaryKeyConstraintEnabled",
                        "0x3662": "SecondaryKey",
                        "0x3663": "BigFunnelPOIContentFlags",
                        "0x3664": "ControlDataForInferenceClutterCleanUpAssistant",
                        "0x3665": "BigFunnelMailboxPOIVersion",
                        "0x3666": "BigFunnelMessageUncompressedPOIVersion",
                        "0x3669": "ControlDataForInferenceTimeModelAssistant",
                        "0x366a": "MailCategorizerProcessedVersion",
                        "0x368e": "BigFunnelPOI",
                        "0x368f": "ContentAggregationFlags",
                        "0x36bf": "UMFaxFolderEntryId",
                        "0x36cc": "RecoveredPublicFolderOriginalPath",
                        "0x36cd": "PerMailboxRecoveryContainerEntryId",
                        "0x36ce": "LostAndFoundFolderEntryId",
                        "0x36cf": "CurrentIPMWasteBasketContainerEntryId",
                        "0x36d6": "RemindersSearchOfflineFolderEntryId",
                        "0x36db": "ContainerTimestamp",
                        "0x36dc": "AppointmentColorName",
                        "0x36dd": "INetUnread",
                        "0x36de": "NetFolderFlags",
                        "0x36df": "FolderWebViewInfo",
                        "0x36e0": "FolderWebViewInfoExtended",
                        "0x36e1": "FolderViewFlags",
                        "0x36e6": "DefaultPostDisplayName",
                        "0x36eb": "FolderViewList",
                        "0x36ec": "AgingPeriod",
                        "0x36ee": "AgingGranularity",
                        "0x36f0": "DefaultFoldersLocaleId",
                        "0x36f1": "InternalAccess",
                        "0x36f2": "PublicFolderSplitStateBinary",
                        "0x36f3": "PublicFolderHierarchySyncNotificationsFolderEntryId",
                        "0x36f4": "IncludeInContentIndex",
                        "0x36f5": "PublicFolderProcessorsStateBinary",
                        "0x36f6": "SystemUse",
                        "0x36f7": "LowLatencyContainerId",
                        "0x36f8": "LowLatencyContainerQuota",
                        "0x3710": "AttachmentMimeSequence",
                        "0x371c": "FailedInboundICalAsAttachment",
                        "0x3720": "AttachmentOriginalUrl",
                        "0x3880": "SyncEventSuppressGuid",
                        "0x39fd": "ListOfContactPhoneNumbersAndEmails",
                        "0x3a3f": "SkipSynchronousDelivery",
                        "0x3a76": "PartnerNetworkId",
                        "0x3a77": "PartnerNetworkUserId",
                        "0x3a78": "PartnerNetworkThumbnailPhotoUrl",
                        "0x3a79": "PartnerNetworkProfilePhotoUrl",
                        "0x3a7a": "PartnerNetworkContactType",
                        "0x3a7b": "RelevanceScore",
                        "0x3a7c": "IsDistributionListContact",
                        "0x3a7d": "IsPromotedContact",
                        "0x3bfa": "UserConfiguredConnectedAccounts",
                        "0x3bfe": "OrgUnitName",
                        "0x3bff": "OrganizationName",
                        "0x3d0b": "ServiceEntryName",
                        "0x3d22": "Win32NTSecurityDescriptor",
                        "0x3d23": "NonWin32ACL",
                        "0x3d24": "ItemLevelACL",
                        "0x3d2e": "ICSGid",
                        "0x3d87": "BigFunnelPostingReplicationScavengedBucketsWatermark",
                        "0x3d88": "BigFunnelPostingListReplicationScavengedBucketsAllowed",
                        "0x3d89": "BigFunnelLastCleanupMaintenance",
                        "0x3d8b": "BigFunnelPostingListLastCompactionMerge",
                        "0x3d8c": "BigFunnelAttributeVectorSharePointDataV1",
                        "0x3d8d": "BigFunnelL1PropertyLengths2V1",
                        "0x3d8e": "BigFunnelL1PropertyLengths1V1Rebuild",
                        "0x3d8f": "MessageSubmittedByOutlook",
                        "0x3d90": "BigFunnelPostingListTargetTableVersion",
                        "0x3d91": "BigFunnelPostingListTargetTableChunkSize",
                        "0x3d92": "BigFunnelL1PropertyLengths1V1",
                        "0x3d93": "ScopeKeyTokenType",
                        "0x3d94": "BigFunnelPostingListTableVersion",
                        "0x3d95": "BigFunnelPostingListTableChunkSize",
                        "0x3d96": "LastTableSizeStatisticsUpdate",
                        "0x3d97": "IcsRestrictionMatch",
                        "0x3d98": "BigFunnelPartialPOI",
                        "0x3d99": "LargeReservedDocIdRanges",
                        "0x3d9a": "DocIdAsImmutableIdGuid",
                        "0x3d9b": "MoveCompletionTime",
                        "0x3d9c": "MaterializedRestrictionSearchRoot",
                        "0x3d9d": "ScheduledISIntegCorruptionCount",
                        "0x3d9e": "ScheduledISIntegExecutionTime",
                        "0x3da1": "QueryCriteriaInternal",
                        "0x3da2": "LastQuotaNotificationTime",
                        "0x3da3": "PropertyPromotionInProgressHiddenItems",
                        "0x3da4": "PropertyPromotionInProgressNormalItems",
                        "0x3da5": "VirtualParentDisplay",
                        "0x3da6": "MailboxTypeDetail",
                        "0x3da7": "InternalTenantHint",
                        "0x3da8": "InternalConversationIndexTracking",
                        "0x3da9": "InternalConversationIndex",
                        "0x3daa": "ConversationItemConversationId",
                        "0x3dab": "VirtualUnreadMessageCount",
                        "0x3dac": "VirtualIsRead",
                        "0x3dad": "IsReadColumn",
                        "0x3dae": "PersistableTenantPartitionHint",
                        "0x3daf": "Internal9ByteChangeNumber",
                        "0x3db0": "Internal9ByteReadCnNew",
                        "0x3db1": "CategoryHeaderLevelStub1",
                        "0x3db2": "CategoryHeaderLevelStub2",
                        "0x3db3": "CategoryHeaderLevelStub3",
                        "0x3db4": "CategoryHeaderAggregateProp0",
                        "0x3db5": "CategoryHeaderAggregateProp1",
                        "0x3db6": "CategoryHeaderAggregateProp2",
                        "0x3db7": "CategoryHeaderAggregateProp3",
                        "0x3db8": "MailboxMoveExtendedFlags",
                        "0x3dbb": "MaintenanceId",
                        "0x3dbc": "MailboxType",
                        "0x3dbd": "MessageFlagsActual",
                        "0x3dbe": "InternalChangeKey",
                        "0x3dbf": "InternalSourceKey",
                        "0x3dd1": "CorrelationId",
                        "0x3e00": "IdentityDisplay",
                        "0x3e01": "IdentityEntryId",
                        "0x3e02": "ResourceMethods",
                        "0x3e03": "ResourceType",
                        "0x3e04": "StatusCode",
                        "0x3e05": "IdentitySearchKey",
                        "0x3e06": "OwnStoreEntryId",
                        "0x3e07": "ResourcePath",
                        "0x3e08": "StatusString",
                        "0x3e0b": "RemoteProgress",
                        "0x3e0c": "RemoteProgressText",
                        "0x3e0d": "RemoteValidateOK",
                        "0x3f00": "ControlFlags",
                        "0x3f01": "ControlStructure",
                        "0x3f02": "ControlType",
                        "0x3f03": "DeltaX",
                        "0x3f04": "DeltaY",
                        "0x3f05": "XPos",
                        "0x3f06": "YPos",
                        "0x3f07": "ControlId",
                        "0x3f88": "AttachmentId",
                        "0x3f89": "GVid",
                        "0x3f8a": "GDID",
                        "0x3f95": "XVid",
                        "0x3f96": "GDefVid",
                        "0x3fc8": "ReplicaChangeNumber",
                        "0x3fc9": "LastConflict",
                        "0x3fd4": "RMI",
                        "0x3fd5": "InternalPostReply",
                        "0x3fd6": "NTSDModificationTime",
                        "0x3fd7": "ACLDataChecksum",
                        "0x3fd8": "PreviewUnread",
                        "0x3fe4": "DesignInProgress",
                        "0x3fe5": "SecureOrigination",
                        "0x3fe8": "AddressBookDisplayName",
                        "0x3ff2": "RuleTriggerHistory",
                        "0x3ff3": "MoveToStoreEntryId",
                        "0x3ff4": "MoveToFolderEntryId",
                        "0x3ffe": "QuotaType",
                        "0x4000": "NewAttachment",
                        "0x4001": "StartEmbed",
                        "0x4002": "EndEmbed",
                        "0x4003": "StartRecip",
                        "0x4004": "EndRecip",
                        "0x4005": "EndCcRecip",
                        "0x4006": "EndBccRecip",
                        "0x4007": "EndP1Recip",
                        "0x4008": "DNPrefix",
                        "0x4009": "StartTopFolder",
                        "0x400a": "StartSubFolder",
                        "0x400b": "EndFolder",
                        "0x400c": "StartMessage",
                        "0x400d": "EndMessage",
                        "0x400e": "EndAttachment",
                        "0x400f": "EcWarning",
                        "0x4010": "StartFAIMessage",
                        "0x4011": "NewFXFolder",
                        "0x4012": "IncrSyncChange",
                        "0x4013": "IncrSyncDelete",
                        "0x4014": "IncrSyncEnd",
                        "0x4015": "IncrSyncMessage",
                        "0x4016": "FastTransferDelProp",
                        "0x4017": "IdsetGiven",
                        "0x4018": "FastTransferErrorInfo",
                        "0x4019": "SenderFlags",
                        "0x401b": "ReceivedByFlags",
                        "0x401c": "ReceivedRepresentingFlags",
                        "0x401d": "OriginalSenderFlags",
                        "0x401e": "OriginalSentRepresentingFlags",
                        "0x401f": "ReportFlags",
                        "0x4020": "ReadReceiptFlags",
                        "0x4021": "SoftDeletes",
                        "0x4022": "CreatorAddressType",
                        "0x4023": "CreatorEmailAddress",
                        "0x4024": "LastModifierAddressType",
                        "0x4025": "LastModifierEmailAddress",
                        "0x4026": "ReportAddressType",
                        "0x4027": "ReportEmailAddress",
                        "0x4028": "ReportDisplayName",
                        "0x402d": "IdsetRead",
                        "0x402e": "IdsetUnread",
                        "0x402f": "IncrSyncRead",
                        "0x4037": "ReportSimpleDisplayName",
                        "0x4038": "CreatorSimpleDisplayName",
                        "0x4039": "LastModifierSimpleDisplayName",
                        "0x403a": "IncrSyncStateBegin",
                        "0x403b": "IncrSyncStateEnd",
                        "0x403c": "IncrSyncImailStream",
                        "0x403f": "SenderOriginalAddressType",
                        "0x4040": "SenderOriginalEmailAddress",
                        "0x4041": "SentRepresentingOriginalAddressType",
                        "0x4042": "SentRepresentingOriginalEmailAddress",
                        "0x4043": "OriginalSenderOriginalAddressType",
                        "0x4044": "OriginalSenderOriginalEmailAddress",
                        "0x4045": "OriginalSentRepresentingOriginalAddressType",
                        "0x4046": "OriginalSentRepresentingOriginalEmailAddress",
                        "0x4047": "ReceivedByOriginalAddressType",
                        "0x4048": "ReceivedByOriginalEmailAddress",
                        "0x4049": "ReceivedRepresentingOriginalAddressType",
                        "0x404a": "ReceivedRepresentingOriginalEmailAddress",
                        "0x404b": "ReadReceiptOriginalAddressType",
                        "0x404c": "ReadReceiptOriginalEmailAddress",
                        "0x404d": "ReportOriginalAddressType",
                        "0x404e": "ReportOriginalEmailAddress",
                        "0x404f": "CreatorOriginalAddressType",
                        "0x4050": "CreatorOriginalEmailAddress",
                        "0x4051": "LastModifierOriginalAddressType",
                        "0x4052": "LastModifierOriginalEmailAddress",
                        "0x4053": "OriginatorOriginalAddressType",
                        "0x4054": "OriginatorOriginalEmailAddress",
                        "0x4055": "ReportDestinationOriginalAddressType",
                        "0x4056": "ReportDestinationOriginalEmailAddress",
                        "0x4057": "OriginalAuthorOriginalAddressType",
                        "0x4058": "OriginalAuthorOriginalEmailAddress",
                        "0x4059": "CreatorFlags",
                        "0x405a": "LastModifierFlags",
                        "0x405b": "OriginatorFlags",
                        "0x405c": "ReportDestinationFlags",
                        "0x405d": "OriginalAuthorFlags",
                        "0x405e": "OriginatorSimpleDisplayName",
                        "0x405f": "ReportDestinationSimpleDisplayName",
                        "0x4061": "OriginatorSearchKey",
                        "0x4062": "ReportDestinationAddressType",
                        "0x4063": "ReportDestinationEmailAddress",
                        "0x4064": "ReportDestinationSearchKey",
                        "0x4066": "IncrSyncImailStreamContinue",
                        "0x4067": "IncrSyncImailStreamCancel",
                        "0x4071": "IncrSyncImailStream2Continue",
                        "0x4074": "IncrSyncProgressMode",
                        "0x4075": "SyncProgressPerMsg",
                        "0x407a": "IncrSyncMsgPartial",
                        "0x407b": "IncrSyncGroupInfo",
                        "0x407c": "IncrSyncGroupId",
                        "0x407d": "IncrSyncChangePartial",
                        "0x4084": "ContentFilterPCL",
                        "0x4085": "PeopleInsightsLastAccessTime",
                        "0x4086": "EmailUsageLastActivityTime",
                        "0x4087": "PdpProfileDataMigrationFlags",
                        "0x4088": "ControlDataForPdpDataMigrationAssistant",
                        "0x5500": "IsInterestingForFileExtraction",
                        "0x5d03": "OriginalSenderSMTPAddress",
                        "0x5d04": "OriginalSentRepresentingSMTPAddress",
                        "0x5d06": "OriginalAuthorSMTPAddress",
                        "0x5d09": "MessageUsageData",
                        "0x5d0a": "CreatorSMTPAddress",
                        "0x5d0b": "LastModifierSMTPAddress",
                        "0x5d0c": "ReportSMTPAddress",
                        "0x5d0d": "OriginatorSMTPAddress",
                        "0x5d0e": "ReportDestinationSMTPAddress",
                        "0x5fe5": "RecipientSipUri",
                        "0x6000": "RssServerLockStartTime",
                        "0x6001": "DotStuffState",
                        "0x6002": "RssServerLockClientName",
                        "0x61af": "ScheduleData",
                        "0x65ef": "RuleMsgActions",
                        "0x65f0": "RuleMsgCondition",
                        "0x65f1": "RuleMsgConditionLCID",
                        "0x65f4": "PreventMsgCreate",
                        "0x65f9": "LISSD",
                        "0x65fa": "IMAPUnsubscribeList",
                        "0x6607": "ProfileUnresolvedName",
                        "0x660d": "ProfileMaxRestrict",
                        "0x660e": "ProfileABFilesPath",
                        "0x660f": "ProfileFavFolderDisplayName",
                        "0x6613": "ProfileHomeServerAddrs",
                        "0x662b": "TestLineSpeed",
                        "0x6630": "LegacyShortcutsFolderEntryId",
                        "0x6635": "FavoritesDefaultName",
                        "0x6641": "DeletedFolderCount",
                        "0x6643": "DeletedAssociatedMessageCount32",
                        "0x6644": "ReplicaServer",
                        "0x664c": "FidMid",
                        "0x6652": "ActiveUserEntryId",
                        "0x6655": "ICSChangeKey",
                        "0x6657": "SetPropsCondition",
                        "0x6659": "InternetContent",
                        "0x665b": "OriginatorName",
                        "0x665c": "OriginatorEmailAddress",
                        "0x665d": "OriginatorAddressType",
                        "0x665e": "OriginatorEntryId",
                        "0x6662": "RecipientNumber",
                        "0x6664": "ReportDestinationName",
                        "0x6665": "ReportDestinationEntryId",
                        "0x6692": "ReplicationMsgPriority",
                        "0x6697": "WorkerProcessId",
                        "0x669a": "CurrentDatabaseSchemaVersion",
                        "0x669e": "SecureInSite",
                        "0x66a7": "MailboxFlags",
                        "0x66ab": "MailboxMessagesPerFolderCountWarningQuota",
                        "0x66ac": "MailboxMessagesPerFolderCountReceiveQuota",
                        "0x66ad": "NormalMessagesWithAttachmentsCount32",
                        "0x66ae": "AssociatedMessagesWithAttachmentsCount32",
                        "0x66af": "FolderHierarchyChildrenCountWarningQuota",
                        "0x66b0": "FolderHierarchyChildrenCountReceiveQuota",
                        "0x66b1": "AttachmentsOnNormalMessagesCount32",
                        "0x66b2": "AttachmentsOnAssociatedMessagesCount32",
                        "0x66b3": "NormalMessageSize64",
                        "0x66e0": "ServerDN",
                        "0x66e1": "BackfillRanking",
                        "0x66e2": "LastTransmissionTime",
                        "0x66e3": "StatusSendTime",
                        "0x66e4": "BackfillEntryCount",
                        "0x66e5": "NextBroadcastTime",
                        "0x66e6": "NextBackfillTime",
                        "0x66e7": "LastCNBroadcast",
                        "0x66eb": "BackfillId",
                        "0x66f4": "LastShortCNBroadcast",
                        "0x66fb": "AverageTransmissionTime",
                        "0x66fc": "ReplicationStatus",
                        "0x66fd": "LastDataReceivalTime",
                        "0x670c": "AutoReset",
                        "0x6712": "ScopeFIDs",
                        "0x671e": "PFPlatinumHomeMdb",
                        "0x673f": "ReadCnNewExport",
                        "0x6745": "LocallyDelivered",
                        "0x6746": "MimeSize",
                        "0x6747": "FileSize",
                        "0x674b": "CategID",
                        "0x674c": "ParentCategID",
                        "0x6750": "ChangeType",
                        "0x6753": "Not822Renderable",
                        "0x6758": "LTID",
                        "0x6759": "CnExport",
                        "0x675a": "PclExport",
                        "0x675b": "CnMvExport",
                        "0x675c": "MidsetDeletedExport",
                        "0x675d": "ArticleNumMic",
                        "0x675e": "ArticleNumMost",
                        "0x6760": "RulesSync",
                        "0x6762": "ReplicaListRC",
                        "0x6763": "ReplicaListRBUG",
                        "0x6766": "FIDC",
                        "0x676b": "MailboxOwnerDN",
                        "0x676d": "IMAPCachedBodystructure",
                        "0x676f": "AltRecipientDN",
                        "0x6770": "NoLocalDelivery",
                        "0x6771": "DeliveryContentLength",
                        "0x6772": "AutoReply",
                        "0x6773": "MailboxOwnerDisplayName",
                        "0x6774": "MailboxLastUpdated",
                        "0x6775": "AdminSurName",
                        "0x6776": "AdminGivenName",
                        "0x6777": "ActiveSearchCount",
                        "0x677a": "OverQuotaLimit",
                        "0x677c": "SubmitContentLength",
                        "0x677d": "LogonRightsOnMailbox",
                        "0x677e": "ReservedIdCounterRangeUpperLimit",
                        "0x677f": "ReservedCnCounterRangeUpperLimit",
                        "0x6780": "SetReceiveCount",
                        "0x6781": "BigFunnelIsEnabled",
                        "0x6782": "SubmittedCount",
                        "0x6783": "CreatorToken",
                        "0x6785": "SearchFIDs",
                        "0x6786": "RecursiveSearchFIDs",
                        "0x678a": "CategFIDs",
                        "0x6791": "MidSegmentStart",
                        "0x6792": "MidsetDeleted",
                        "0x6793": "MidsetExpired",
                        "0x6794": "CnsetIn",
                        "0x6796": "CnsetBackfill",
                        "0x6798": "MidsetTombstones",
                        "0x679a": "GWFolder",
                        "0x679b": "IPMFolder",
                        "0x679c": "PublicFolderPath",
                        "0x679f": "MidSegmentIndex",
                        "0x67a0": "MidSegmentSize",
                        "0x67a1": "CnSegmentStart",
                        "0x67a2": "CnSegmentIndex",
                        "0x67a3": "CnSegmentSize",
                        "0x67a5": "PCL",
                        "0x67a6": "CnMv",
                        "0x67a7": "FolderTreeRootFID",
                        "0x67a8": "SourceEntryId",
                        "0x67a9": "MailFlags",
                        "0x67ab": "SubmitResponsibility",
                        "0x67ad": "SharedReceiptHandling",
                        "0x67b3": "MessageAttachmentList",
                        "0x67b5": "SenderCAI",
                        "0x67b6": "SentRepresentingCAI",
                        "0x67b7": "OriginalSenderCAI",
                        "0x67b8": "OriginalSentRepresentingCAI",
                        "0x67b9": "ReceivedByCAI",
                        "0x67ba": "ReceivedRepresentingCAI",
                        "0x67bb": "ReadReceiptCAI",
                        "0x67bc": "ReportCAI",
                        "0x67bd": "CreatorCAI",
                        "0x67be": "LastModifierCAI",
                        "0x67c4": "AnonymousRights",
                        "0x67ce": "SearchGUID",
                        "0x67d2": "CnsetRead",
                        "0x67da": "CnsetBackfillFAI",
                        "0x67de": "ReplMsgVersion",
                        "0x67e5": "IdSetDeleted",
                        "0x67e6": "FolderMessages",
                        "0x67e7": "SenderReplid",
                        "0x67e8": "CnMin",
                        "0x67e9": "CnMax",
                        "0x67ea": "ReplMsgType",
                        "0x67eb": "RgszDNResponders",
                        "0x67f2": "ViewCoveringPropertyTags",
                        "0x67f4": "ICSViewFilter",
                        "0x67f8": "OriginatorCAI",
                        "0x67f9": "ReportDestinationCAI",
                        "0x67fa": "OriginalAuthorCAI",
                        "0x6807": "EventCounter",
                        "0x6809": "EventFid",
                        "0x680a": "EventMid",
                        "0x680b": "EventFidParent",
                        "0x680c": "EventFidOld",
                        "0x680e": "EventFidOldParent",
                        "0x680f": "EventCreatedTime",
                        "0x6811": "EventItemCount",
                        "0x6812": "EventFidRoot",
                        "0x6818": "EventExtendedFlags",
                        "0x681c": "EventImmutableid",
                        "0x681d": "MailboxQuarantineEnd",
                        "0x681e": "EventOldParentDefaultFolderType",
                        "0x681f": "MailboxNumber",
                        "0x6821": "InferenceClientId",
                        "0x6822": "InferenceItemId",
                        "0x6823": "InferenceCreateTime",
                        "0x6824": "InferenceWindowId",
                        "0x6825": "InferenceSessionId",
                        "0x6826": "InferenceFolderId",
                        "0x682e": "InferenceTimeZone",
                        "0x682f": "InferenceCategory",
                        "0x6832": "InferenceModuleSelected",
                        "0x6833": "InferenceLayoutType",
                        "0x6835": "InferenceTimeStamp",
                        "0x6836": "InferenceOLKUserActivityLoggingEnabled",
                        "0x6837": "InferenceClientVersion",
                        "0x6838": "InferenceSSISource",
                        "0x6839": "ActivityWorkload",
                        "0x683b": "ActivityItemType",
                        "0x683c": "ActivityContainerMailbox",
                        "0x683d": "ActivityContainerId",
                        "0x683e": "ActivityNonExoItemId",
                        "0x683f": "ActivityClientInstanceId",
                        "0x6840": "ActivityImmutableItemId",
                        "0x6857": "AgingAgeFolder",
                        "0x6858": "AgingDontAgeMe",
                        "0x6859": "AgingFileNameAfter9",
                        "0x685b": "AgingWhenDeletedOnServer",
                        "0x685c": "AgingWaitUntilExpired",
                        "0x685f": "ActivityImmutableEntryId",
                        "0x6870": "DelegateEntryIds2",
                        "0x6871": "DelegateFlags2",
                        "0x6873": "InferenceTrainedModelVersionBreadCrumb",
                        "0x6874": "FolderPathFullName",
                        "0x6875": "ImmutableIdExport",
                        "0x6876": "ControlDataForTrendingAroundMeAssistant",
                        "0x6877": "RestrictionAnnotationWordBreakingTokens",
                        "0x6878": "RestrictionAnnotationWordBreakingTokenLengths",
                        "0x6879": "ControlDataForCalendarInsightsAssistant",
                        "0x687a": "ControlDataForFreeBusyPublishingTimeBasedAssistant",
                        "0x687b": "RestrictionAnnotationIndexPropertyTag",
                        "0x687c": "IsAbandonedMoveDestination",
                        "0x687d": "ImmutableId26Bytes",
                        "0x687e": "ImmutableIdSetIn",
                        "0x687f": "SearchFolderLargeRestriction",
                        "0x689f": "ConversationMsgSentTime",
                        "0x68a4": "PersonCompanyNameMailboxWide",
                        "0x68a5": "PersonDisplayNameMailboxWide",
                        "0x68a6": "PersonGivenNameMailboxWide",
                        "0x68a7": "PersonSurnameMailboxWide",
                        "0x68a8": "PersonPhotoContactEntryIdMailboxWide",
                        "0x68b0": "PersonFileAsMailboxWide",
                        "0x68b1": "PersonRelevanceScoreMailboxWide",
                        "0x68b2": "PersonIsDistributionListMailboxWide",
                        "0x68b3": "PersonHomeCityMailboxWide",
                        "0x68b4": "PersonCreationTimeMailboxWide",
                        "0x68b7": "PersonGALLinkIDMailboxWide",
                        "0x68ba": "PersonMvEmailAddressMailboxWide",
                        "0x68bb": "PersonMvEmailDisplayNameMailboxWide",
                        "0x68bc": "PersonMvEmailRoutingTypeMailboxWide",
                        "0x68bd": "PersonImAddressMailboxWide",
                        "0x68be": "PersonWorkCityMailboxWide",
                        "0x68bf": "PersonDisplayNameFirstLastMailboxWide",
                        "0x68c0": "PersonDisplayNameLastFirstMailboxWide",
                        "0x68c2": "ConversationHasClutter",
                        "0x68c3": "ConversationHasClutterMailboxWide",
                        "0x68c4": "ExchangeObjectId",
                        "0x68c5": "ViewLargeRestriction",
                        "0x68c6": "ClientDiagnosticLevel",
                        "0x68c7": "ClientDiagnosticData",
                        "0x6900": "ConversationLastMemberDocumentId",
                        "0x6901": "ConversationPreview",
                        "0x6902": "ConversationLastMemberDocumentIdMailboxWide",
                        "0x6903": "ConversationInitialMemberDocumentId",
                        "0x6904": "ConversationMemberDocumentIds",
                        "0x6905": "ConversationMessageDeliveryOrRenewTimeMailboxWide",
                        "0x6907": "ConversationMessageRichContentMailboxWide",
                        "0x6908": "ConversationPreviewMailboxWide",
                        "0x6909": "ConversationMessageDeliveryOrRenewTime",
                        "0x690a": "ConversationWorkingSetSourcePartition",
                        "0x690b": "ConversationSystemCategories",
                        "0x690c": "ConversationSystemCategoriesMailboxWide",
                        "0x690d": "ConversationExchangeApplicationFlagsMailboxWide",
                        "0x690e": "ConversationMvMentionsMailboxWide",
                        "0x690f": "ConversationMvMentions",
                        "0x6910": "ConversationMvThreadIds",
                        "0x6911": "ConversationMvThreadIdsMailboxWide",
                        "0x6912": "ConversationLikeCountMailboxWide",
                        "0x6913": "ConversationReturnTime",
                        "0x6914": "ConversationReturnTimeMailboxWide",
                        "0x6915": "UserActivityPayloadVersion",
                        "0x6916": "ConversationAtAllMention",
                        "0x6917": "ConversationAtAllMentionMailboxWide",
                        "0x6918": "ConversationInferenceClassification",
                        "0x6919": "ConversationCharm",
                        "0x691a": "ConversationCharmMailboxWide",
                        "0x691b": "SignalTypeId",
                        "0x691c": "SignalAppId",
                        "0x691d": "SignalActorId",
                        "0x691e": "SignalClientVersion",
                        "0x691f": "SignalAadTenantId",
                        "0x6920": "SignalActorIdType",
                        "0x6921": "SignalOS",
                        "0x6922": "SignalOSVersion",
                        "0x6923": "SignalLatitude",
                        "0x6924": "SignalCv",
                        "0x6925": "SignalClientIp",
                        "0x6926": "SignalUserAgent",
                        "0x6927": "SignalDeviceId",
                        "0x6928": "SignalSchemaVersion",
                        "0x6929": "SignalAppWorkload",
                        "0x692a": "SignalCompliance",
                        "0x692b": "SignalItemType",
                        "0x692c": "SignalContainerId",
                        "0x692d": "SignalContainerType",
                        "0x692e": "SignalLongitude",
                        "0x6930": "SignalLocationType",
                        "0x6931": "SignalPrecision",
                        "0x6932": "SignalLocaleId",
                        "0x6933": "SignalTargetItemId",
                        "0x6935": "SignalTimeStamp",
                        "0x6938": "SignalIsClient",
                        "0x693a": "ConversationSenderName",
                        "0x693b": "ConversationSenderNameMailboxWide",
                        "0x693c": "ConversationSenderSmtpAddress",
                        "0x693d": "ConversationSenderSmtpAddressMailboxWide",
                        "0x693e": "ConversationMemberCnSet",
                        "0x693f": "ConversationMemberCnSetMailboxWide",
                        "0x6940": "ConversationMemberImmutableIdSet",
                        "0x6941": "ConversationMemberImmutableIdSetMailboxWide",
                        "0x6942": "ConversationLastAttachmentsProcessedTime",
                        "0x6943": "ConversationLastAttachmentsProcessedTimeMailboxWide",
                        "0x694a": "SignalAppName",
                        "0x694b": "SignalTimeStampOffset",
                        "0x6e01": "SecurityFlags",
                        "0x6e04": "SecurityReceiptRequestProcessed",
                        "0x7000": "UserInformationInstanceCreationTime",
                        "0x7018": "RemoteFolderDisplayName",
                        "0x7019": "AssistantFilterResult",
                        "0x701a": "ControlDataForTimerBrokerAssistant",
                        "0x701b": "ContainsScheduledTimers",
                        "0x701c": "LastScheduledTimerChangeToken",
                        "0x701d": "ControlDataForBigFunnelStoreIndexAssistant",
                        "0x701e": "BigFunnelStoreIndexAssistantProcessedVersion",
                        "0x701f": "BigFunnelStoreIndexAssistantRequestedVersion",
                        "0x7020": "PeopleRelevanceMailEventCount",
                        "0x7021": "PeopleRelevanceCalendarEventCount",
                        "0x7022": "PeopleRelevanceContactEventCount",
                        "0x7023": "PeopleRelevanceLastSuccessfulRunTime",
                        "0x702a": "MailboxFeatureStorageProperty",
                        "0x702b": "LastUserActionTime",
                        "0x7c00": "FavoriteDisplayName",
                        "0x7c02": "FavPublicSourceKey",
                        "0x7c03": "SyncFolderSourceKey",
                        "0x7c04": "SyncFolderChangeKey",
                        "0x7c09": "UserConfigurationStream",
                        "0x7c0a": "SyncStateBlob",
                        "0x7c0b": "ReplyForwardStatus",
                        "0x7c0c": "PopImapPoisonMessageStamp",
                        "0x7c14": "ControlDataForSystemCleanupFolderAssistant",
                        "0x7c1a": "UserPhotoCacheId",
                        "0x7c1b": "UserPhotoPreviewCacheId",
                        "0x7c1c": "ProfileHeaderPhotoCacheId",
                        "0x7c1d": "ProfileHeaderPhotoPreviewCacheId",
                        "0x7d03": "FavLevelMask",
                        "0x7d0e": "ImmutableId",
                        "0x7d0f": "ControlDataForWeveMessageAssistant",
                        "0x7d10": "WeveMessageAssistantLastMessageSentTime",
                        "0x7d11": "WeveMessageAssistantLastNotificationMessageSentTime",
                        "0x7d12": "ControlDataForMailboxLifecycleAssistant",
                        "0x7d13": "WeveMessageAssistantLastLicenseCheckTime",
                        "0x7d14": "WeveMessageAssistantLicenseExists",
                        "0x7d15": "ReplacedImmutableIdBin",
                        "0x7d16": "IsCloudCacheCrawlingComplete",
                        "0x7d17": "CloudCacheItemSyncStatus",
                        "0x7d18": "ControlDataForRecordReviewAssistant",
                        "0x7d19": "LastRecordIdentifiedTime",
                        "0x7d20": "ControlDataForPeopleRelevanceMultiStepAssistant",
                        "0x7d21": "ControlDataForXrmSharingMaintenanceAssistant",
                        "0x7d23": "ItemAssistantCrawlVersionBlob",
                        "0x7d24": "ControlDataForDynamicTba0",
                        "0x7d25": "ControlDataForDynamicTba1",
                        "0x7d26": "ControlDataForDynamicTba2",
                        "0x7d27": "ControlDataForDynamicTba3",
                        "0x7d28": "ControlDataForDynamicTba4",
                        "0x7d29": "ControlDataForDynamicGriffinTba0",
                        "0x7d2a": "ControlDataForDynamicGriffinTba1",
                        "0x7d2b": "ControlDataForDynamicGriffinTba2",
                        "0x7d2c": "ControlDataForDynamicGriffinTba3",
                        "0x7d2d": "ControlDataForDynamicGriffinTba4",
                        "0x7d2e": "ControlDataForDynamicGriffinTba5",
                        "0x7d2f": "ControlDataForDynamicGriffinTba6",
                        "0x7d30": "ControlDataForDynamicGriffinTba7",
                        "0x7d31": "ControlDataForDynamicGriffinTba8",
                        "0x7d32": "ControlDataForDynamicGriffinTba9",
                        "0x7d33": "ControlDataForDynamicGriffinTba10",
                        "0x7d34": "ControlDataForDynamicGriffinTba11",
                        "0x7d35": "ControlDataForDynamicGriffinTba12",
                        "0x7d36": "ControlDataForDynamicGriffinTba13",
                        "0x7d37": "ControlDataForDynamicGriffinTba14",
                        "0x7d38": "ControlDataForDynamicGriffinTba15",
                        "0x7d39": "ControlDataForDynamicGriffinTba16",
                        "0x7d3a": "ControlDataForDynamicGriffinTba17",
                        "0x7d3b": "ControlDataForDynamicGriffinTba18",
                        "0x7d3c": "ControlDataForDynamicGriffinTba19",
                        "0x7d3d": "ControlDataForDynamicGriffinTba20",
                        "0x7d3e": "ControlDataForDynamicGriffinTba21",
                        "0x7d3f": "ControlDataForDynamicGriffinTba22",
                        "0x7d40": "ControlDataForDynamicGriffinTba23",
                        "0x7d41": "ControlDataForDynamicGriffinTba24",
                        "0x7d42": "ControlDataForDynamicGriffinTba25",
                        "0x7d43": "ControlDataForDynamicGriffinTba26",
                        "0x7d44": "ControlDataForDynamicGriffinTba27",
                        "0x7d45": "ControlDataForDynamicGriffinTba28",
                        "0x7d46": "ControlDataForDynamicGriffinTba29",
                        "0x7d47": "ControlDataForDynamicGriffinTba30",
                        "0x7d48": "ControlDataForContentClassificationAssistant",
                        "0x7d49": "ControlDataForOfficeGraphSecondaryCopyQuotaTimeBasedAssistant",
                        "0x7d4a": "ControlDataForOfficeGraphConvertFilesToWorkingSetAndSpoolsSubFoldersTimeBasedAssistant",
                        "0x7d4b": "ControlDataForOfficeGraphSpoolsScaleOutTimeBasedAssistant",
                        "0x7d7f": "ControlDataForDynamicTba6",
                        "0x7d80": "MailboxTenantSizeEstimate",
                        "0x7d81": "ControlDataForMailboxTenantDataAssistant",
                        "0x7ff7": "IsATPEncrypted",
                        "0x7ff8": "HasDlpDetectedAttachmentClassifications",
                        "0x8d0d": "ExternalDirectoryObjectId"
                    },

                    CLASS_MAPPING: {
                        ATTACHMENT_DATA: '3701'
                    },
                    TYPE_MAPPING: {
                        '001e': 'string',
                        '001f': 'unicode',
                        '0102': 'binary'
                    },
                    DIR_TYPE: {
                        INNER_MSG: '000d'
                    }
                }
            }
        };
    }

    // extractor structure to manage bat/sbat block types and different data types
    static get extractorFieldValue() {
        return {
            sbat: {
                'extractor': function extractDataViaSbat(ds, msgData, fieldProperty, dataTypeExtractor) {
                    let chain = MsgReader.#getChainByBlockSmall(ds, msgData, fieldProperty);
                    if (chain.length === 1) {
                      return MsgReader.#readDataByBlockSmall(ds, msgData, fieldProperty.startBlock, fieldProperty.sizeBlock, dataTypeExtractor);
                    } else if (chain.length > 1) {
                      return MsgReader.#readChainDataByBlockSmall(ds, msgData, fieldProperty, chain, dataTypeExtractor);
                    }
                    return null;
                },
                dataType: {
                    'string':
                        function extractBatString(ds, msgData, blockStartOffset, bigBlockOffset, blockSize) {
                            ds.seek(blockStartOffset + bigBlockOffset);
                            return ds.readString(blockSize);
                        },
                    'unicode':
                        function extractBatUnicode(ds, msgData, blockStartOffset, bigBlockOffset, blockSize) {
                            ds.seek(blockStartOffset + bigBlockOffset);
                            return ds.readUCS2String(blockSize / 2);
                        },
                    'binary':
                        function extractBatBinary(ds, msgData, blockStartOffset, bigBlockOffset, blockSize) {
                            ds.seek(blockStartOffset + bigBlockOffset);
                            return ds.readUint8Array(blockSize);
                        }
                }
            },
            bat: {
                'extractor':
                    function extractDataViaBat(ds, msgData, fieldProperty, dataTypeExtractor) {
                        let offset = MsgReader.#getBlockOffsetAt(msgData, fieldProperty.startBlock);
                        ds.seek(offset);
                        return dataTypeExtractor(ds, fieldProperty);
                    },
                dataType: {
                    'string': function extractSbatString(ds, fieldProperty) {
                        return ds.readString(fieldProperty.sizeBlock);
                    },
                    'unicode': function extractSbatUnicode(ds, fieldProperty) {
                        return ds.readUCS2String(fieldProperty.sizeBlock / 2);
                    },
                    'binary': function extractSbatBinary(ds, fieldProperty) {
                        return ds.readUint8Array(fieldProperty.sizeBlock);
                    }
                }
            }
        };
    }


    // ----------------------------
    // PRIVATE FUNCTIONS
    // ----------------------------

    #decodeRfc1342(string) {
        // =?utf-8?Q?Kostensch=C3=A4tzung=5F451.pdf?=
        const decoder = new TextDecoder();
        string = string.replace(/=\?([0-9a-z\-_:]+)\?(B|Q)\?(.*?)\?=/ig, (m, charset, encoding, encodedText) => {
            let buf = null;
            switch (encoding.toUpperCase()) {
                case 'B': buf = this.#decodeBase64(encodedText, charset); break;
                case 'Q': buf = this.#decodeQuotedPrintable(encodedText, charset, true); break;
                default: throw new Error('invalid string encoding "' + encoding + '"');
            }
            return decoder.decode(new Uint8Array(buf));
        });

        return string;
    }

    /**
     * @param {Uint8Array|String} raw
     * @param {String|null} charset
     * @returns {ArrayBuffer}
     */
    #decodeBase64(raw, charset=null) {
        if (raw instanceof Uint8Array) {
            const decoder = new TextDecoder();
            raw = decoder.decode(raw);
        }
        const binary_string = window.atob(raw);
        const len = binary_string.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binary_string.charCodeAt(i);
        }
        if (!charset) {
            return bytes.buffer;

        } else {
            // convert to utf-8
            const dec = new TextDecoder(charset), enc = new TextEncoder();
            const arr = enc.encode(dec.decode(bytes));
            return arr.buffer;
        }
    }

    /**
     * @param {Uint8Array|String} raw
     * @param {String} charset
     * @param {Bool} replaceUnderline
     * @returns {ArrayBuffer}
     */
    #decodeQuotedPrintable(raw, charset, replaceUnderline=false) {
        if (raw instanceof Uint8Array) {
            const decoder = new TextDecoder();
            raw = decoder.decode(raw);
        }

        // in RFC 1342 underline is used for space
        if (replaceUnderline) {
            raw = raw.replace(/_/g, ' ');
        }

        const dc = new TextDecoder(charset ? charset : 'utf-8');
        const str = raw.replace(/[\t\x20]$/gm, "").replace(/=(?:\r\n?|\n)/g, "").replace(/((?:=[a-fA-F0-9]{2})+)/g, (m) => {
            const cd = m.substring(1).split('='), uArr=new Uint8Array(cd.length);
            for (let i = 0; i < cd.length; i++) {
                uArr[i] = parseInt(cd[i], 16);
            }
            return dc.decode(uArr);
        });

        const encoder = new TextEncoder();
        const arr = encoder.encode(str);
        return arr.buffer;
    }


    // ----------------------------
    // PRIVATE STATIC FUNCTIONS
    // ----------------------------


    static #splitHeaders(headerRaw) {
        const headers = headerRaw.split(/\n(?=[^\s])/g), responseHeaders = {};

        for (let header of headers) {
            const sepPos = header.indexOf(':');
            if (sepPos !== -1) {
                const key = header.substring(0, sepPos).toLowerCase().trim(), value=header.substring(sepPos+1).trim();

                if (responseHeaders[key] && typeof responseHeaders[key] === 'string') {
                    responseHeaders[key] = [responseHeaders[key]];
                }
                if (responseHeaders[key]) {
                    responseHeaders[key].push(value);

                } else {
                    responseHeaders[key] = value;
                }
            }
        }

        return responseHeaders;
    }

    static #getEmailIfValid(str) {
        const validRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/;

        if(str.match(validRegex)) {
            return str;
        } else {
            return null;
        }
    }

    // unit utils
    static #arraysEqual(a, b) {
        if (a === b)
            return true;
        if (!a || !b)
            return false;
        if (a.length !== b.length)
            return false;

        for (let i = 0; i < a.length; i++) {
            if (a[i] !== b[i])
                return false;
        }
        return true;
    }

    static #uInt2int(data) {
        let result = new Array(data.length);
        for (let i = 0; i < data.length; i++) {
            result[i] = data[i] << 24 >> 24;
        }
        return result;
    }

    // MSG Reader implementation

    // check MSG file header
    static #isMSGFile(ds) {
        ds.seek(0);
        return MsgReader.#arraysEqual(MsgReader.CONST.FILE_HEADER, ds.readInt8Array(MsgReader.CONST.FILE_HEADER.length));
    }

    // FAT utils
    static #getBlockOffsetAt(msgData, offset) {
        return (offset + 1) * msgData.bigBlockSize;
    }

    static #getBlockAt(ds, msgData, offset) {
        let startOffset = MsgReader.#getBlockOffsetAt(msgData, offset);
        ds.seek(startOffset);
        return ds.readInt32Array(msgData.bigBlockLength);
    }

    static #getNextBlockInner(ds, msgData, offset, blockOffsetData) {
        let currentBlock = Math.floor(offset / msgData.bigBlockLength);
        let currentBlockIndex = offset % msgData.bigBlockLength;

        let startBlockOffset = blockOffsetData[currentBlock];

        return MsgReader.#getBlockAt(ds, msgData, startBlockOffset)[currentBlockIndex];
    }

    static #getNextBlock(ds, msgData, offset) {
        return MsgReader.#getNextBlockInner(ds, msgData, offset, msgData.batData);
    }

    static #getNextBlockSmall(ds, msgData, offset) {
        return MsgReader.#getNextBlockInner(ds, msgData, offset, msgData.sbatData);
    }

    // convert binary data to dictionary
    static #parseMsgData(ds) {
        let msgData = MsgReader.#headerData(ds);
        msgData.batData = MsgReader.#batData(ds, msgData);
        msgData.sbatData = MsgReader.#sbatData(ds, msgData);
        if (msgData.xbatCount > 0) {
            MsgReader.#xbatData(ds, msgData);
        }
        msgData.propertyData = MsgReader.#propertyData(ds, msgData);
        msgData.fieldsData = MsgReader.#fieldsData(ds, msgData);

        return msgData;
    }

    // extract header data
    static #headerData(ds) {
        let headerData = {};

        // system data
        headerData.bigBlockSize =
                ds.readByte(/*const position*/30) === MsgReader.CONST.MSG.L_BIG_BLOCK_MARK ? MsgReader.CONST.MSG.L_BIG_BLOCK_SIZE : MsgReader.CONST.MSG.S_BIG_BLOCK_SIZE;
        headerData.bigBlockLength = headerData.bigBlockSize / 4;
        headerData.xBlockLength = headerData.bigBlockLength - 1;

        // header data
        headerData.batCount = ds.readInt(MsgReader.CONST.MSG.HEADER.BAT_COUNT_OFFSET);
        headerData.propertyStart = ds.readInt(MsgReader.CONST.MSG.HEADER.PROPERTY_START_OFFSET);
        headerData.sbatStart = ds.readInt(MsgReader.CONST.MSG.HEADER.SBAT_START_OFFSET);
        headerData.sbatCount = ds.readInt(MsgReader.CONST.MSG.HEADER.SBAT_COUNT_OFFSET);
        headerData.xbatStart = ds.readInt(MsgReader.CONST.MSG.HEADER.XBAT_START_OFFSET);
        headerData.xbatCount = ds.readInt(MsgReader.CONST.MSG.HEADER.XBAT_COUNT_OFFSET);

        return headerData;
    }

    static #batCountInHeader(msgData) {
        let maxBatsInHeader = (MsgReader.CONST.MSG.S_BIG_BLOCK_SIZE - MsgReader.CONST.MSG.HEADER.BAT_START_OFFSET) / 4;
        return Math.min(msgData.batCount, maxBatsInHeader);
    }

    static #batData(ds, msgData) {
        let result = new Array(MsgReader.#batCountInHeader(msgData));
        ds.seek(MsgReader.CONST.MSG.HEADER.BAT_START_OFFSET);
        for (let i = 0; i < result.length; i++) {
            result[i] = ds.readInt32();
        }
        return result;
    }

    static #sbatData(ds, msgData) {
        let result = [];
        let startIndex = msgData.sbatStart;

        for (let i = 0; i < msgData.sbatCount && startIndex !== MsgReader.CONST.MSG.END_OF_CHAIN; i++) {
            result.push(startIndex);
            startIndex = MsgReader.#getNextBlock(ds, msgData, startIndex);
        }
        return result;
    }

    static #xbatData(ds, msgData) {
        let batCount = MsgReader.#batCountInHeader(msgData);
        let batCountTotal = msgData.batCount;
        let remainingBlocks = batCountTotal - batCount;

        let nextBlockAt = msgData.xbatStart;
        for (let i = 0; i < msgData.xbatCount; i++) {
            let xBatBlock = MsgReader.#getBlockAt(ds, msgData, nextBlockAt);
            nextBlockAt = xBatBlock[msgData.xBlockLength];

            let blocksToProcess = Math.min(remainingBlocks, msgData.xBlockLength);
            for (let j = 0; j < blocksToProcess; j++) {
                let blockStartAt = xBatBlock[j];
                if (blockStartAt === MsgReader.CONST.MSG.UNUSED_BLOCK || blockStartAt === MsgReader.CONST.MSG.END_OF_CHAIN) {
                    break;
                }
                msgData.batData.push(blockStartAt);
            }
            remainingBlocks -= blocksToProcess;
        }
    }

    // extract property data and property hierarchy
    static #propertyData(ds, msgData) {
        let props = [];

        let currentOffset = msgData.propertyStart;

        while (currentOffset !== MsgReader.CONST.MSG.END_OF_CHAIN) {
            MsgReader.#convertBlockToProperties(ds, msgData, currentOffset, props);
            currentOffset = MsgReader.#getNextBlock(ds, msgData, currentOffset);
        }
        MsgReader.#createPropertyHierarchy(props, /*property with index 0 (zero) always as root*/props[0]);
        return props;
    }

    static #convertName(ds, offset) {
        let nameLength = ds.readShort(offset + MsgReader.CONST.MSG.PROP.NAME_SIZE_OFFSET);
        if (nameLength < 1) {
            return '';
        } else {
            return ds.readStringAt(offset, nameLength / 2).replace(/\0/g, '');;
        }
    }

    static #convertProperty(ds, index, offset) {
        return {
            index: index,
            type: ds.readByte(offset + MsgReader.CONST.MSG.PROP.TYPE_OFFSET),
            name: MsgReader.#convertName(ds, offset),
            // hierarchy
            previousProperty: ds.readInt(offset + MsgReader.CONST.MSG.PROP.PREVIOUS_PROPERTY_OFFSET),
            nextProperty: ds.readInt(offset + MsgReader.CONST.MSG.PROP.NEXT_PROPERTY_OFFSET),
            childProperty: ds.readInt(offset + MsgReader.CONST.MSG.PROP.CHILD_PROPERTY_OFFSET),
            // data offset
            startBlock: ds.readInt(offset + MsgReader.CONST.MSG.PROP.START_BLOCK_OFFSET),
            sizeBlock: ds.readInt(offset + MsgReader.CONST.MSG.PROP.SIZE_OFFSET)
        };
    }

    static #convertBlockToProperties(ds, msgData, propertyBlockOffset, props) {

        let propertyCount = msgData.bigBlockSize / MsgReader.CONST.MSG.PROP.PROPERTY_SIZE;
        let propertyOffset = MsgReader.#getBlockOffsetAt(msgData, propertyBlockOffset);

        for (let i = 0; i < propertyCount; i++) {
            let propertyType = ds.readByte(propertyOffset + MsgReader.CONST.MSG.PROP.TYPE_OFFSET);
            switch (propertyType) {
                case MsgReader.CONST.MSG.PROP.TYPE_ENUM.ROOT:
                case MsgReader.CONST.MSG.PROP.TYPE_ENUM.DIRECTORY:
                case MsgReader.CONST.MSG.PROP.TYPE_ENUM.DOCUMENT:
                    props.push(MsgReader.#convertProperty(ds, props.length, propertyOffset));
                    break;
                default:
                    /* unknown property types */
                    props.push(null);
            }

            propertyOffset += MsgReader.CONST.MSG.PROP.PROPERTY_SIZE;
        }
    }

    static #createPropertyHierarchy(props, nodeProperty) {

        if (nodeProperty.childProperty === MsgReader.CONST.MSG.PROP.NO_INDEX) {
            return;
        }
        nodeProperty.children = [];

        let children = [nodeProperty.childProperty];
        while (children.length !== 0) {
            let currentIndex = children.shift();
            let current = props[currentIndex];
            if (!current) {
                continue;
            }
            nodeProperty.children.push(currentIndex);

            if (current.type === MsgReader.CONST.MSG.PROP.TYPE_ENUM.DIRECTORY) {
                MsgReader.#createPropertyHierarchy(props, current);
            }
            if (current.previousProperty !== MsgReader.CONST.MSG.PROP.NO_INDEX) {
                children.push(current.previousProperty);
            }
            if (current.nextProperty !== MsgReader.CONST.MSG.PROP.NO_INDEX) {
                children.push(current.nextProperty);
            }
        }
    }

    // extract real fields
    static #fieldsData(ds, msgData) {
        let fields = {
            attachments: [],
            recipients: []
        };
        MsgReader.#fieldsDataDir(ds, msgData, msgData.propertyData[0], fields);
        return fields;
    }

    static #fieldsDataDir(ds, msgData, dirProperty, fields) {
        const dirIsRoot = dirProperty && dirProperty.name && dirProperty.name.toLowerCase().indexOf("root") !== -1;

        if (dirProperty.children && dirProperty.children.length > 0) {
            for (let i = 0; i < dirProperty.children.length; i++) {
                let childProperty = msgData.propertyData[dirProperty.children[i]];

                if (childProperty.type === MsgReader.CONST.MSG.PROP.TYPE_ENUM.DIRECTORY) {
                    MsgReader.#fieldsDataDirInner(ds, msgData, childProperty, fields);

                } else if (childProperty.type === MsgReader.CONST.MSG.PROP.TYPE_ENUM.DOCUMENT && childProperty.name.indexOf(MsgReader.CONST.MSG.FIELD.PREFIX.DOCUMENT) === 0) {
                    MsgReader.#fieldsDataDocument(ds, msgData, childProperty, fields);

                // root properties
                } else if (dirIsRoot && childProperty.name === '__properties_version1.0') {
                    let binPropertiesData = MsgReader.#getFieldValue(ds, msgData, childProperty, 'binary');
                    fields._properties = MsgReader.#readRootProperties(binPropertiesData);
                }
            }
        }
    }

    /**
     * read the Property Stream
     * https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg/20c1125f-043d-42d9-b1dc-cb9b7e5198ef
     */
    static #readRootProperties(binPropertiesData) {
        const ds = new DataStream(binPropertiesData, 0, DataStream.LITTLE_ENDIAN),
                values={
                    _header:{}
                };

        // go to 0
        ds.seek(0);

        // Reserved (8 bytes): This field MUST be set to zero
        for (let i = 0; i < 8; i++) {
            if (ds.readUint8() !== 0) {
                throw new Error('invalid properties header');
            }
        }

        values._header.nextRecipientId = ds.readUint32();
        values._header.nextAttachmentId = ds.readUint32();
        values._header.recipientCount = ds.readUint32();
        values._header.attachmentCount = ds.readUint32();

        // Reserved (8 bytes): This field MUST be set to zero
        for (let i = 0; i < 8; i++) {
            if (ds.readUint8() !== 0) {
                throw new Error('invalid properties header');
            }
        }

        // The data inside the property stream MUST be an array of 16-byte entries
        while (!ds.isEof()) {
            const property = {flags: {}, type: null, binData: null, data: null};

            let fieldType = ds.readUint16(),
                    fieldClass = ds.readUint16(),
                    flags = ds.readUint32().toString(16),
                    value = [],
                    fieldName = MsgReader.#getMapiFieldName(fieldClass.toString(16).padStart(4, '0'));

            property.type = fieldType;
            property.typeStr = fieldType.toString(16).padStart(4, '0');;
            property.flags.mandatory = !!(flags & 0x00000001);
            property.flags.readable = !!(flags & 0x00000002);
            property.flags.writeable = !!(flags & 0x00000004);

            for (let i = 0; i < 8; i++) {
                value.push(ds.readUint8());
            }

            if (!fieldName) {
                fieldName = '_' + property.typeStr;
            }

            property.binData = new Uint8Array(value);
            property.data = MsgReader.#readFieldData(property.typeStr, property.binData);

            // write to object
            values[fieldName] = property;
        }

        return values;

    }

    /**
     * read the field data
     * @param {Number|String} ftype
     * @param {Uint8Array} binaryData
     * @returns {Mixed}
     */
    static #readFieldData(ftype, binaryData) {
        if (typeof ftype === 'number') {
            ftype = ftype.toString(16).padStart(4, '0');
        }

        switch (ftype.toLowerCase()) {
            case '0040': return MsgReader.#convertFiletimeToDateTime(binaryData);
            case '0001': return null;
            case '0002': return MsgReader.#uint8ArrayToInt(binaryData, 16, true); // Integer 16-bit signed
            case '0003': return MsgReader.#uint8ArrayToInt(binaryData, 32, true); // Integer 32-bit signed
            case '0014': return MsgReader.#uint8ArrayToInt(binaryData, 64, true); // Integer 64-bit signed
            case '000b': return MsgReader.#uint8ArrayToInt(binaryData, 64, true) !== 0; // Boolean
        }

        return null;
    }

    static #fieldsDataDirInner(ds, msgData, dirProperty, fields) {
        if (dirProperty.name.indexOf(MsgReader.CONST.MSG.FIELD.PREFIX.ATTACHMENT) === 0) {

            // attachment
            let attachmentField = {};
            fields.attachments.push(attachmentField);
            MsgReader.#fieldsDataDir(ds, msgData, dirProperty, attachmentField);
        } else if (dirProperty.name.indexOf(MsgReader.CONST.MSG.FIELD.PREFIX.RECIPIENT) === 0) {

            // recipient
            let recipientField = {};
            fields.recipients.push(recipientField);
            MsgReader.#fieldsDataDir(ds, msgData, dirProperty, recipientField);
        } else {

            // other dir
            let childFieldType = MsgReader.#getFieldType(dirProperty);
            if (childFieldType !== MsgReader.CONST.MSG.FIELD.DIR_TYPE.INNER_MSG) {
                MsgReader.#fieldsDataDir(ds, msgData, dirProperty, fields);
            } else {
                // MSG as attachment currently isn't supported
                fields.innerMsgContent = true;
            }
        }
    }

    static #isAddPropertyValue(fieldName, fieldTypeMapped) {
        return fieldName !== 'Body' || fieldTypeMapped !== 'binary';
    }

    static #fieldsDataDocument(ds, msgData, documentProperty, fields) {
        let value = documentProperty.name.substring(12).toLowerCase();
        let fieldClass = value.substring(0, 4);
        let fieldType = value.substring(4, 8);
        let fieldName = MsgReader.#getMapiFieldName(fieldClass) ?? 'unknown_'+fieldClass;
        let suffix='', offset=1;

        while (fields[fieldName + suffix]) {
            offset++;
            suffix = '-' + offset;
        }

        fieldName = fieldName + suffix;

        //let fieldName = MsgReader.CONST.MSG.FIELD.NAME_MAPPING[fieldClass];
        let fieldTypeMapped = MsgReader.CONST.MSG.FIELD.TYPE_MAPPING[fieldType];

        if (fieldName) {
            let fieldValue = MsgReader.#getFieldValue(ds, msgData, documentProperty, fieldTypeMapped ?? 'binary');

            if (MsgReader.#isAddPropertyValue(fieldName, fieldTypeMapped)) {
                fields[fieldName] = MsgReader.#applyValueConverter(fieldName, fieldTypeMapped, fieldValue);
            }

        }

        // Attachment handling
        if (fieldClass === MsgReader.CONST.MSG.FIELD.CLASS_MAPPING.ATTACHMENT_DATA) {

            // attachment specific info
            fields['dataId'] = documentProperty.index;
            fields['contentLength'] = documentProperty.sizeBlock;
        }
    }

    /**
     * converts a windows timestamp to a Date object
     * @param {uint8Array} filetime
     * @returns {Date}
     */
    static #convertFiletimeToDateTime(filetime) {
        // Extract low and high parts from the Uint8Array
        const lowPart = filetime[0] |
          (filetime[1] << 8) |
          (filetime[2] << 16) |
          (filetime[3] << 24);
        const highPart = filetime[4] |
          (filetime[5] << 8) |
          (filetime[6] << 16) |
          (filetime[7] << 24);

        // Combine low and high parts to get the full 64-bit value
        const filetimeValue = (BigInt(highPart) << BigInt(32)) | BigInt(lowPart);

        // Windows Filetime starts from January 1, 1601 (in 100-nanosecond intervals)
        const ticksPerMillisecond = BigInt(10000);
        const ticksPerSecond = BigInt(10000000);
        const ticksPerDay = ticksPerSecond * BigInt(60) * BigInt(60) * BigInt(24);
        const epochTicks = BigInt(116444736000000000);

        // Calculate the number of ticks since the Unix epoch (January 1, 1970)
        const unixTicks = (filetimeValue - epochTicks) / ticksPerMillisecond;

        // Create a JavaScript Date object from the Unix ticks
        const date = new Date(Number(unixTicks));

        return date;
    }

    /**
     * get infos about a mapi field
     */
    static #getMapiFieldName(fieldClass) {
        return MsgReader.CONST.MSG.FIELD.MAPI_PROPERTIES['0x' + fieldClass.toLowerCase()] ?? null;
    }

    static #applyValueConverter(fieldName, fieldTypeMapped, fieldValue) {
        if (fieldTypeMapped === 'binary' && fieldName === 'BodyHtml') {
            return MsgReader.#convertUint8ArrayToString(fieldValue);
        }
        return fieldValue;
    }

    static #getFieldType(fieldProperty) {
        let value = fieldProperty.name.substring(12).toLowerCase();
        return value.substring(4, 8);
    }

    static #readDataByBlockSmall(ds, msgData, startBlock, blockSize, dataTypeExtractor) {
        let byteOffset = startBlock * MsgReader.CONST.MSG.SMALL_BLOCK_SIZE;
        let bigBlockNumber = Math.floor(byteOffset / msgData.bigBlockSize);
        let bigBlockOffset = byteOffset % msgData.bigBlockSize;

        let rootProp = msgData.propertyData[0];

        let nextBlock = rootProp.startBlock;
        for (let i = 0; i < bigBlockNumber; i++) {
            nextBlock = MsgReader.#getNextBlock(ds, msgData, nextBlock);
        }
        let blockStartOffset = MsgReader.#getBlockOffsetAt(msgData, nextBlock);

        return dataTypeExtractor(ds, msgData, blockStartOffset, bigBlockOffset, blockSize);
    }

    static #readChainDataByBlockSmall(ds, msgData, fieldProperty, chain, dataTypeExtractor) {
        let resultData = new Int8Array(fieldProperty.sizeBlock);

        for (let i = 0, idx = 0; i < chain.length; i++) {
            let data = MsgReader.#readDataByBlockSmall(ds, msgData, chain[i], MsgReader.CONST.MSG.SMALL_BLOCK_SIZE, MsgReader.extractorFieldValue.sbat.dataType.binary);
            for (let j = 0; j < data.length; j++) {
                resultData[idx++] = data[j];
            }
        }
        let localDs = new DataStream(resultData, 0, DataStream.LITTLE_ENDIAN);
        return dataTypeExtractor(localDs, msgData, 0, 0, fieldProperty.sizeBlock);
    }

    static #getChainByBlockSmall(ds, msgData, fieldProperty) {
        let blockChain = [];
        let nextBlockSmall = fieldProperty.startBlock;
        while (nextBlockSmall !== MsgReader.CONST.MSG.END_OF_CHAIN) {
            blockChain.push(nextBlockSmall);
            nextBlockSmall = MsgReader.#getNextBlockSmall(ds, msgData, nextBlockSmall);
        }
        return blockChain;
    }

    static #getFieldValue(ds, msgData, fieldProperty, typeMapped) {
        let value = null;

        let valueExtractor = fieldProperty.sizeBlock < MsgReader.CONST.MSG.BIG_BLOCK_MIN_DOC_SIZE ? MsgReader.extractorFieldValue.sbat : MsgReader.extractorFieldValue.bat;
        let dataTypeExtractor = valueExtractor.dataType[typeMapped];

        if (dataTypeExtractor) {
            value = valueExtractor.extractor(ds, msgData, fieldProperty, dataTypeExtractor);
        }
        return value;
    }

    static #convertUint8ArrayToString(uint8ArraValue) {
        return new TextDecoder('utf-8').decode(uint8ArraValue);
    }

    static #uint8ArrayToInt(uint8Array, size, signed) {
        if (size % 8 !== 0 || size < 8 || size > 64) {
          throw new Error('Invalid size. Must be 8, 16, 32 or 64.');
        }

        const buffer = new ArrayBuffer(size/8);
        const dataView = new DataView(buffer);

        // Copy the values from the Uint8Array into the buffer
        for (let i = 0; i < (size/8); i++) {
            dataView.setUint8(i, uint8Array[i]);
        }

        // Retrieve the integer value based on the specified signedness
        let value;
        switch (size) {
            case 8:
                value = signed ? dataView.getInt8(0) : dataView.getUint8(0);
                break;
            case 16:
                value = signed ? dataView.getInt16(0) : dataView.getUint16(0);
                break;
            case 32:
                value = signed ? dataView.getInt32(0) : dataView.getUint32(0);
                break;
            case 64:
                if (signed) {
                    const high = dataView.getInt32(0);
                    const low = dataView.getUint32(4);
                    value = BigInt(low) + (BigInt(high) << BigInt(32));
                } else {
                    const high = dataView.getUint32(0);
                    const low = dataView.getUint32(4);
                    value = BigInt(low) + (BigInt(high) << BigInt(32));
                }
                break;
        }

        return value;
    }
}
