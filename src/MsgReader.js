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

        if (this.#fileData.fieldsData && this.#fileData.fieldsData.headers) {
            this.#headers = MsgReader.#splitHeaders(this.#fileData.fieldsData.headers);
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

        return {fileName: attachData.fileName, content: fieldData};
    }

    getAttachments() {
        let attachments=[];

        if (this.#fileData.fieldsData && this.#fileData.fieldsData.attachments && this.#fileData.fieldsData.attachments.length > 0) {
            for (const atm of this.#fileData.fieldsData.attachments) {
                let attachment = this.getAttachment(atm);

                    attachments.push({
                        filename: attachment.fileName,
                        contentType: atm.mimeType,
                        content: attachment.content,
                        filesize: atm.contentLength,
                        pidContentId: atm.pidContentId
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
        return null;
    }

    getSubject() {
        const hSub = this.getHeader('subject', true, true);
        if (hSub) {
            return hSub;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.subject) {
            return this.#fileData.fieldsData.subject;
        }

    }

    getFrom() {
        const hSub = this.getHeader('from', true, true);
        if (hSub) {
            return hSub;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.senderName && this.#fileData.fieldsData.senderMail) {
            return this.#fileData.fieldsData.senderName + ' <' + this.#fileData.fieldsData.senderMail + '>';

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.senderMail) {
            return this.#fileData.fieldsData.senderMail;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.senderName) {
            return this.#fileData.fieldsData.senderName;
        }

        return '';
    }

    getCc() {
        return this.getHeader('cc', true, true);
    }

    getTo() {
        const hSub = this.getHeader('to', true, true);
        if (hSub) {
            return hSub;

        } else if (this.#fileData.fieldsData && this.#fileData.fieldsData.recipients && this.#fileData.fieldsData.recipients.length > 0) {
            let rep = '';
            for (const rm of this.#fileData.fieldsData.recipients) {
                if (rm.name && rm.email) {
                    if (rep) {
                        rep += '; ';
                    }
                    rep += rm.name + ' <' + rm.email + '>';

                } else if (rm.email) {
                    if (rep) {
                        rep += '; ';
                    }
                    rep += rm.email;

                } else if (rm.name) {
                    if (rep) {
                        rep += '; ';
                    }
                    rep += rm.name;
                }
            }
            return rep;
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
        return this.#fileData.fieldsData.body;
    }

    getMessageHtml() {
        return this.#fileData.fieldsData.bodyHTML;
    }

    // ----------------------------
    // PRIVATE STATIC CONSTANTS
    // ----------------------------

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
                        '0037': 'subject',
                        '0c1a': 'senderName',
                        '5d02': 'senderEmail',
                        '1000': 'body',
                        '1013': 'bodyHTML',
                        '007d': 'headers',

                        // attachment specific
                        '3703': 'extension',
                        '3704': 'fileNameShort',
                        '3707': 'fileName',
                        '3712': 'pidContentId',
                        '370e': 'mimeType',

                        // recipient specific
                        '3001': 'name',
                        '39fe': 'email'
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
            return ds.readStringAt(offset, nameLength / 2);
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

        if (dirProperty.children && dirProperty.children.length > 0) {
            for (let i = 0; i < dirProperty.children.length; i++) {
                let childProperty = msgData.propertyData[dirProperty.children[i]];

                if (childProperty.type === MsgReader.CONST.MSG.PROP.TYPE_ENUM.DIRECTORY) {
                    MsgReader.#fieldsDataDirInner(ds, msgData, childProperty, fields);

                } else if (childProperty.type === MsgReader.CONST.MSG.PROP.TYPE_ENUM.DOCUMENT && childProperty.name.indexOf(MsgReader.CONST.MSG.FIELD.PREFIX.DOCUMENT) === 0) {
                    MsgReader.#fieldsDataDocument(ds, msgData, childProperty, fields);
                }
            }
        }
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
        return fieldName !== 'body' || fieldTypeMapped !== 'binary';
    }

    static #fieldsDataDocument(ds, msgData, documentProperty, fields) {
        let value = documentProperty.name.substring(12).toLowerCase();
        let fieldClass = value.substring(0, 4);
        let fieldType = value.substring(4, 8);

        let fieldName = MsgReader.CONST.MSG.FIELD.NAME_MAPPING[fieldClass];
        let fieldTypeMapped = MsgReader.CONST.MSG.FIELD.TYPE_MAPPING[fieldType];

        if (fieldName) {
            let fieldValue = MsgReader.#getFieldValue(ds, msgData, documentProperty, fieldTypeMapped);

            if (MsgReader.#isAddPropertyValue(fieldName, fieldTypeMapped)) {
                fields[fieldName] = MsgReader.#applyValueConverter(fieldName, fieldTypeMapped, fieldValue);
            }
        }
        if (fieldClass === MsgReader.CONST.MSG.FIELD.CLASS_MAPPING.ATTACHMENT_DATA) {

            // attachment specific info
            fields['dataId'] = documentProperty.index;
            fields['contentLength'] = documentProperty.sizeBlock;
        }
    }

    // todo: html body test
    static #applyValueConverter(fieldName, fieldTypeMapped, fieldValue) {
        if (fieldTypeMapped === 'binary' && fieldName === 'bodyHTML') {
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
}