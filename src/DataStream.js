/*
 * Copyright 2016  Ilmari Heikkinen
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

/**
 DataStream reads scalars, arrays and structs of data from an ArrayBuffer.
 It's like a file-like DataView on steroids.
 */
export class DataStream {
    #buffer;
    #dataView;
    #byteOffset = 0;

    /**
     Virtual byte length of the DataStream backing buffer.
     Updated to be max of original buffer size and last written size.
     If dynamicSize is false is set to buffer size.
     @type {number}
     */
    #byteLength = 0;

    /**
     Whether to extend DataStream buffer when trying to write beyond its size.
     If set, the buffer is reallocated to twice its current size until the
     requested write fits the buffer.
     @type {boolean}
     */
    #dynamicSize = true;

    /**
     Seek position where DataStream#readStruct ran into a problem.
     Useful for debugging struct parsing.

     @type {number}
     */
    failurePosition = 0;

    /**
     Big-endian const to use as default endianness.
     @type {boolean}
     */
    static BIG_ENDIAN = false;

    /**
     Little-endian const to use as default endianness.
     @type {boolean}
     */
    static LITTLE_ENDIAN = true;

    /**
     Native endianness. Either DataStream.BIG_ENDIAN or DataStream.LITTLE_ENDIAN
     depending on the platform endianness.

     @type {boolean}
     */
    static endianness = new Int8Array(new Int16Array([1]).buffer)[0] > 0;

    // ----------------------------
    // CONSTRUCTOR
    // ----------------------------

    /**
     * @param {ArrayBuffer} arrayBuffer ArrayBuffer to read from.
     * @param {?Number} byteOffset Offset from arrayBuffer beginning for the DataStream.
     * @param {?Boolean} endianness DataStream.BIG_ENDIAN or DataStream.LITTLE_ENDIAN (the default).
     * @returns {undefined}
     */
    constructor(arrayBuffer, byteOffset, endianness) {
        this.#byteOffset = byteOffset || 0;

        if (arrayBuffer instanceof ArrayBuffer) {
            this.buffer = arrayBuffer;

        } else if (typeof arrayBuffer == "object") {
            this.dataView = arrayBuffer;

            if (byteOffset) {
                this.#byteOffset += byteOffset;
            }

        } else {
            this.buffer = new ArrayBuffer(arrayBuffer || 1);
        }

        this.position = 0;
        this.endianness = endianness == null ? DataStream.LITTLE_ENDIAN : endianness;
    }

    // ----------------------------
    // GETTER SETTER
    // ----------------------------

    get dynamicSize() {
        return this.#dynamicSize;
    }
    set dynamicSize(v) {
        if (!v) {
            this.#trimAlloc();
        }
        this.#dynamicSize = v;
    }

    /**
     Returns the byte length of the DataStream object.
     @type {number}
     */
    get byteLength() {
        return this.#byteLength - this.#byteOffset;
    }

    /**
     Set/get the backing ArrayBuffer of the DataStream object.
     The setter updates the DataView to point to the new buffer.
     @type {Object}
     */
    get buffer() {
        this.#trimAlloc();
        return this.#buffer;
    }
    set buffer(v) {
        this.#buffer = v;
        this.#dataView = new DataView(this.#buffer, this.#byteOffset);
        this.#byteLength = this.#buffer.byteLength;
    }

    /**
     Set/get the byteOffset of the DataStream object.
     The setter updates the DataView to point to the new byteOffset.
     @type {number}
     */
    get byteOffset() {
        return this.#byteOffset;
    }
    set byteOffset(v) {
        this.#byteOffset = v;
        this.#dataView = new DataView(this.#buffer, this.#byteOffset);
        this.#byteLength = this.#buffer.byteLength;
    }

    /**
     Set/get the backing DataView of the DataStream object.
     The setter updates the buffer and byteOffset to point to the DataView values.
     @type {Object}
     */
    get dataView() {
        return this.#dataView;
    }
    set dataView(v) {
        this.#byteOffset = v.byteOffset;
        this.#buffer = v.buffer;
        this.#dataView = new DataView(this.#buffer, this.#byteOffset);
        this.#byteLength = this.#byteOffset + v.byteLength;
    }

    // ----------------------------
    // PUBLIC FUNCTIONS
    // ----------------------------

    /**
     Saves the DataStream contents to the given filename.
     Uses Chrome's anchor download property to initiate download.

     @param {string} filename Filename to save as.
     @return {null}
     */
    save(filename) {
        let blob = new Blob(this.buffer);
        let URL = (window.webkitURL || window.URL);
        if (URL && URL.createObjectURL) {
            let url = URL.createObjectURL(blob);
            let a = document.createElement('a');
            a.setAttribute('href', url);
            a.setAttribute('download', filename);
            a.click();
            URL.revokeObjectURL(url);
        } else {
            throw("DataStream.save: Can't create object URL.");
        }
    }

    /**
     Internal function to resize the DataStream buffer when required.
     @param {number} extra Number of bytes to add to the buffer allocation.
     @return {null}
     */
    #realloc(extra) {
        if (!this.#dynamicSize) {
            return;
        }
        let req = this.#byteOffset + this.position + extra;
        let blen = this.#buffer.byteLength;
        if (req <= blen) {
            if (req > this.#byteLength) {
                this.#byteLength = req;
            }
            return;
        }
        if (blen < 1) {
            blen = 1;
        }
        while (req > blen) {
            blen *= 2;
        }
        let buf = new ArrayBuffer(blen);
        let src = new Uint8Array(this.#buffer);
        let dst = new Uint8Array(buf, 0, src.length);
        dst.set(src);
        this.buffer = buf;
        this.#byteLength = req;
    }

    /**
     Internal function to trim the DataStream buffer when required.
     Used for stripping out the extra bytes from the backing buffer when
     the virtual byteLength is smaller than the buffer byteLength (happens after
     growing the buffer with writes and not filling the extra space completely).

     @return {null}
     */
    #trimAlloc() {
        if (this.#byteLength == this.#buffer.byteLength) {
            return;
        }
        let buf = new ArrayBuffer(this.#byteLength);
        let dst = new Uint8Array(buf);
        let src = new Uint8Array(this.#buffer, 0, dst.length);
        dst.set(src);
        this.buffer = buf;
    }

    /**
     Sets the DataStream read/write position to given position.
     Clamps between 0 and DataStream length.

     @param {number} pos Position to seek to.
     @return {null}
     */
    seek(pos) {
        let npos = Math.max(0, Math.min(this.byteLength, pos));
        this.position = (isNaN(npos) || !isFinite(npos)) ? 0 : npos;
    }

    /**
     Returns true if the DataStream seek pointer is at the end of buffer and
     there's no more data to read.

     @return {boolean} True if the seek pointer is at the end of the buffer.
     */
    isEof() {
        return (this.position >= this.byteLength);
    }

    /**
     Maps an Int32Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Int32Array to the DataStream backing buffer.
     */
    mapInt32Array(length, e) {
        this.#realloc(length * 4);
        let arr = new Int32Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 4;
        return arr;
    }

    /**
     Maps an Int16Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Int16Array to the DataStream backing buffer.
     */
    mapInt16Array(length, e) {
        this.#realloc(length * 2);
        let arr = new Int16Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 2;
        return arr;
    }

    /**
     Maps an Int8Array into the DataStream buffer.

     Nice for quickly reading in data.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Int8Array to the DataStream backing buffer.
     */
    mapInt8Array(length) {
        this.#realloc(length * 1);
        let arr = new Int8Array(this.#buffer, this.byteOffset + this.position, length);
        this.position += length * 1;
        return arr;
    }

    /**
     Maps a Uint32Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Uint32Array to the DataStream backing buffer.
     */
    mapUint32Array(length, e) {
        this.#realloc(length * 4);
        let arr = new Uint32Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 4;
        return arr;
    }

    /**
     Maps a Uint16Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Uint16Array to the DataStream backing buffer.
     */
    mapUint16Array(length, e) {
        this.#realloc(length * 2);
        let arr = new Uint16Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 2;
        return arr;
    }

    /**
     Maps a Uint8Array into the DataStream buffer.

     Nice for quickly reading in data.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Uint8Array to the DataStream backing buffer.
     */
    mapUint8Array(length) {
        this.#realloc(length * 1);
        let arr = new Uint8Array(this.#buffer, this.byteOffset + this.position, length);
        this.position += length * 1;
        return arr;
    }

    /**
     Maps a Float64Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Float64Array to the DataStream backing buffer.
     */
    mapFloat64Array(length, e) {
        this.#realloc(length * 8);
        let arr = new Float64Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 8;
        return arr;
    }

    /**
     Maps a Float32Array into the DataStream buffer, swizzling it to native
     endianness in-place. The current offset from the start of the buffer needs to
     be a multiple of element size, just like with typed array views.

     Nice for quickly reading in data. Warning: potentially modifies the buffer
     contents.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} Float32Array to the DataStream backing buffer.
     */
    mapFloat32Array(length, e) {
        this.#realloc(length * 4);
        let arr = new Float32Array(this.#buffer, this.byteOffset + this.position, length);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += length * 4;
        return arr;
    }

    /**
     Reads an Int32Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Int32Array.
     */
    readInt32Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 4) : length;
        let arr = new Int32Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads an Int16Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Int16Array.
     */
    readInt16Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 2) : length;
        let arr = new Int16Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads an Int8Array of desired length from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Int8Array.
     */
    readInt8Array(length) {
        length = length == null ? (this.byteLength - this.position) : length;
        let arr = new Int8Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads a Uint32Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Uint32Array.
     */
    readUint32Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 4) : length;
        let arr = new Uint32Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads a Uint16Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Uint16Array.
     */
    readUint16Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 2) : length;
        let arr = new Uint16Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads a Uint8Array of desired length from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Uint8Array.
     */
    readUint8Array(length) {
        length = length == null ? (this.byteLength - this.position) : length;
        let arr = new Uint8Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads a Float64Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Float64Array.
     */
    readFloat64Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 8) : length;
        let arr = new Float64Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Reads a Float32Array of desired length and endianness from the DataStream.

     @param {number} length Number of elements to map.
     @param {?boolean} e Endianness of the data to read.
     @return {Object} The read Float32Array.
     */
    readFloat32Array(length, e) {
        length = length == null ? (this.byteLength - this.position / 4) : length;
        let arr = new Float32Array(length);
        DataStream.memcpy(arr.buffer, 0,
                this.buffer, this.byteOffset + this.position,
                length * arr.BYTES_PER_ELEMENT);
        DataStream.arrayToNative(arr, e == null ? this.endianness : e);
        this.position += arr.byteLength;
        return arr;
    }

    /**
     Writes an Int32Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeInt32Array(arr, e) {
        this.#realloc(arr.length * 4);
        if (arr instanceof Int32Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapInt32Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeInt32(arr[i], e);
            }
        }
    }

    /**
     Writes an Int16Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeInt16Array(arr, e) {
        this.#realloc(arr.length * 2);
        if (arr instanceof Int16Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapInt16Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeInt16(arr[i], e);
            }
        }
    }

    /**
     Writes an Int8Array to the DataStream.

     @param {Object} arr The array to write.
     */
    writeInt8Array(arr) {
        this.#realloc(arr.length * 1);
        if (arr instanceof Int8Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapInt8Array(arr.length);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeInt8(arr[i]);
            }
        }
    }

    /**
     Writes a Uint32Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeUint32Array(arr, e) {
        this.#realloc(arr.length * 4);
        if (arr instanceof Uint32Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapUint32Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeUint32(arr[i], e);
            }
        }
    }

    /**
     Writes a Uint16Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeUint16Array(arr, e) {
        this.#realloc(arr.length * 2);
        if (arr instanceof Uint16Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapUint16Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeUint16(arr[i], e);
            }
        }
    }

    /**
     Writes a Uint8Array to the DataStream.

     @param {Object} arr The array to write.
     */
    writeUint8Array(arr) {
        this.#realloc(arr.length * 1);
        if (arr instanceof Uint8Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapUint8Array(arr.length);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeUint8(arr[i]);
            }
        }
    }

    /**
     Writes a Float64Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeFloat64Array(arr, e) {
        this.#realloc(arr.length * 8);
        if (arr instanceof Float64Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapFloat64Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeFloat64(arr[i], e);
            }
        }
    }

    /**
     Writes a Float32Array of specified endianness to the DataStream.

     @param {Object} arr The array to write.
     @param {?boolean} e Endianness of the data to write.
     */
    writeFloat32Array(arr, e) {
        this.#realloc(arr.length * 4);
        if (arr instanceof Float32Array &&
                this.byteOffset + this.position % arr.BYTES_PER_ELEMENT === 0) {
            DataStream.memcpy(this.#buffer, this.byteOffset + this.position,
                    arr.buffer, 0,
                    arr.byteLength);
            this.mapFloat32Array(arr.length, e);
        } else {
            for (let i = 0; i < arr.length; i++) {
                this.writeFloat32(arr[i], e);
            }
        }
    }

    /**
     Reads a 32-bit int from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readInt32(e) {
        let v = this.#dataView.getInt32(this.position, e == null ? this.endianness : e);
        this.position += 4;
        return v;
    }

    /**
     Reads a 32-bit int from the DataStream with the offset.

     @param {number} offset The offset.
     @return {number} The read number.
     */
    readInt(offset) {
        this.seek(offset);
        return this.readInt32();
    }

    /**
     Reads a 16-bit int from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readInt16(e) {
        let v = this.#dataView.getInt16(this.position, e == null ? this.endianness : e);
        this.position += 2;
        return v;
    }

    /**
     Reads a 16-bit int from the DataStream with the offset

     @param {number} offset The offset.
     @return {number} The read number.
     */
    readShort(offset) {
        this.seek(offset);
        return this.readInt16();
    }

    /**
     Reads an 8-bit int from the DataStream.

     @return {number} The read number.
     */
    readInt8() {
        let v = this.#dataView.getInt8(this.position);
        this.position += 1;
        return v;
    }

    /**
     Reads an 8-bit int from the DataStream with the offset.

     @param {number} offset The offset.
     @return {number} The read number.
     */
    readByte(offset) {
        this.seek(offset);
        return this.readInt8();
    }

    /**
     Reads a 32-bit unsigned int from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readUint32(e) {
        let v = this.#dataView.getUint32(this.position, e == null ? this.endianness : e);
        this.position += 4;
        return v;
    }

    /**
     Reads a 16-bit unsigned int from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readUint16(e) {
        let v = this.#dataView.getUint16(this.position, e == null ? this.endianness : e);
        this.position += 2;
        return v;
    }

    /**
     Reads an 8-bit unsigned int from the DataStream.

     @return {number} The read number.
     */
    readUint8() {
        let v = this.#dataView.getUint8(this.position);
        this.position += 1;
        return v;
    }

    /**
     Reads a 32-bit float from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readFloat32(e) {
        let v = this.#dataView.getFloat32(this.position, e == null ? this.endianness : e);
        this.position += 4;
        return v;
    }

    /**
     Reads a 64-bit float from the DataStream with the desired endianness.

     @param {?boolean} e Endianness of the number.
     @return {number} The read number.
     */
    readFloat64(e) {
        let v = this.#dataView.getFloat64(this.position, e == null ? this.endianness : e);
        this.position += 8;
        return v;
    }

    /**
     Writes a 32-bit int to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeInt32(v, e) {
        this.#realloc(4);
        this.#dataView.setInt32(this.position, v, e == null ? this.endianness : e);
        this.position += 4;
    }

    /**
     Writes a 16-bit int to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeInt16(v, e) {
        this.#realloc(2);
        this.#dataView.setInt16(this.position, v, e == null ? this.endianness : e);
        this.position += 2;
    }

    /**
     Writes an 8-bit int to the DataStream.

     @param {number} v Number to write.
     */
    writeInt8(v) {
        this.#realloc(1);
        this.#dataView.setInt8(this.position, v);
        this.position += 1;
    }

    /**
     Writes a 32-bit unsigned int to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeUint32(v, e) {
        this.#realloc(4);
        this.#dataView.setUint32(this.position, v, e == null ? this.endianness : e);
        this.position += 4;
    }

    /**
     Writes a 16-bit unsigned int to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeUint16(v, e) {
        this.#realloc(2);
        this.#dataView.setUint16(this.position, v, e == null ? this.endianness : e);
        this.position += 2;
    }

    /**
     Writes an 8-bit unsigned  int to the DataStream.

     @param {number} v Number to write.
     */
    writeUint8(v) {
        this.#realloc(1);
        this.#dataView.setUint8(this.position, v);
        this.position += 1;
    }

    /**
     Writes a 32-bit float to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeFloat32(v, e) {
        this.#realloc(4);
        this.#dataView.setFloat32(this.position, v, e == null ? this.endianness : e);
        this.position += 4;
    }

    /**
     Writes a 64-bit float to the DataStream with the desired endianness.

     @param {number} v Number to write.
     @param {?boolean} e Endianness of the number.
     */
    writeFloat64(v, e) {
        this.#realloc(8);
        this.#dataView.setFloat64(this.position, v, e == null ? this.endianness : e);
        this.position += 8;
    }

    /**
     Copies byteLength bytes from the src buffer at srcOffset to the
     dst buffer at dstOffset.

     @param {Object} dst Destination ArrayBuffer to write to.
     @param {number} dstOffset Offset to the destination ArrayBuffer.
     @param {Object} src Source ArrayBuffer to read from.
     @param {number} srcOffset Offset to the source ArrayBuffer.
     @param {number} byteLength Number of bytes to copy.
     */
    static memcpy(dst, dstOffset, src, srcOffset, byteLength) {
        let dstU8 = new Uint8Array(dst, dstOffset, byteLength);
        let srcU8 = new Uint8Array(src, srcOffset, byteLength);
        dstU8.set(srcU8);
    }

    /**
     Converts array to native endianness in-place.

     @param {Object} array Typed array to convert.
     @param {boolean} arrayIsLittleEndian True if the data in the array is
     little-endian. Set false for big-endian.
     @return {Object} The converted typed array.
     */
    static arrayToNative(array, arrayIsLittleEndian) {
        if (arrayIsLittleEndian == this.endianness) {
            return array;
        } else {
            return this.flipArrayEndianness(array);
        }
    }

    /**
     Converts native endianness array to desired endianness in-place.

     @param {Object} array Typed array to convert.
     @param {boolean} littleEndian True if the converted array should be
     little-endian. Set false for big-endian.
     @return {Object} The converted typed array.
     */
    static nativeToEndian(array, littleEndian) {
        if (this.endianness == littleEndian) {
            return array;
        } else {
            return this.flipArrayEndianness(array);
        }
    }

    /**
     Flips typed array endianness in-place.

     @param {Object} array Typed array to flip.
     @return {Object} The converted typed array.
     */
    static flipArrayEndianness(array) {
        let u8 = new Uint8Array(array.buffer, array.byteOffset, array.byteLength);
        for (let i = 0; i < array.byteLength; i += array.BYTES_PER_ELEMENT) {
            for (let j = i + array.BYTES_PER_ELEMENT - 1, k = i; j > k; j--, k++) {
                let tmp = u8[k];
                u8[k] = u8[j];
                u8[j] = tmp;
            }
        }
        return array;
    }

    /**
     Creates an array from an array of character codes.
     Uses String.fromCharCode on the character codes and concats the results into a string.

     @param {array} array Array of character codes.
     @return {string} String created from the character codes.
     **/
    static createStringFromArray(array) {
        let str = "";
        for (let i = 0; i < array.length; i++) {
            str += String.fromCharCode(array[i]);
        }
        return str;
    }

    /**
     Reads a struct of data from the DataStream. The struct is defined as
     a flat array of [name, type]-pairs. See the example below:

     ds.readStruct([
     'headerTag', 'uint32', // Uint32 in DataStream endianness.
     'headerTag2', 'uint32be', // Big-endian Uint32.
     'headerTag3', 'uint32le', // Little-endian Uint32.
     'array', ['[]', 'uint32', 16], // Uint32Array of length 16.
     'array2Length', 'uint32',
     'array2', ['[]', 'uint32', 'array2Length'] // Uint32Array of length array2Length
     ]);

     The possible values for the type are as follows:

     // Number types

     // Unsuffixed number types use DataStream endianness.
     // To explicitly specify endianness, suffix the type with
     // 'le' for little-endian or 'be' for big-endian,
     // e.g. 'int32be' for big-endian int32.

     'uint8' -- 8-bit unsigned int
     'uint16' -- 16-bit unsigned int
     'uint32' -- 32-bit unsigned int
     'int8' -- 8-bit int
     'int16' -- 16-bit int
     'int32' -- 32-bit int
     'float32' -- 32-bit float
     'float64' -- 64-bit float

     // String types
     'cstring' -- ASCII string terminated by a zero byte.
     'string:N' -- ASCII string of length N, where N is a literal integer.
     'string:variableName' -- ASCII string of length $variableName,
     where 'variableName' is a previously parsed number in the current struct.
     'string,CHARSET:N' -- String of byteLength N encoded with given CHARSET.
     'u16string:N' -- UCS-2 string of length N in DataStream endianness.
     'u16stringle:N' -- UCS-2 string of length N in little-endian.
     'u16stringbe:N' -- UCS-2 string of length N in big-endian.

     // Complex types
     [name, type, name_2, type_2, ..., name_N, type_N] -- Struct
     function(dataStream, struct) {} -- Callback function to read and return data.
     {get: function(dataStream, struct) {},
     set: function(dataStream, struct) {}}
     -- Getter/setter functions to read and return data, handy for using the same
     struct definition for reading and writing structs.
     ['[]', type, length] -- Array of given type and length. The length can be either
     a number, a string that references a previously-read
     field, or a callback function(struct, dataStream, type){}.
     If length is '*', reads in as many elements as it can.

     @param {Object} structDefinition Struct definition object.
     @return {Object} The read struct. Null if failed to read struct.
     */
    readStruct(structDefinition) {
        let struct = {}, t, v, n;
        let p = this.position;
        for (let i = 0; i < structDefinition.length; i += 2) {
            t = structDefinition[i + 1];
            v = this.readType(t, struct);
            if (v == null) {
                if (this.failurePosition == 0) {
                    this.failurePosition = this.position;
                }
                this.position = p;
                return null;
            }
            struct[structDefinition[i]] = v;
        }
        return struct;
    }

    /**
     Read UCS-2 string of desired length and endianness from the DataStream.

     @param {number} length The length of the string to read.
     @param {boolean} endianness The endianness of the string data in the DataStream.
     @return {string} The read string.
     */
    readUCS2String(length, endianness) {
        return DataStream.createStringFromArray(this.readUint16Array(length, endianness));
    }

    /**
     Read UCS-2 string of desired length and offset from the DataStream.

     @param {number} offset The offset.
     @param {number} length The length of the string to read.
     @return {string} The read string.
     */
    readStringAt(offset, length) {
        this.seek(offset);
        return this.readUCS2String(length);
    }

    /**
     Write a UCS-2 string of desired endianness to the DataStream. The
     lengthOverride argument lets you define the number of characters to write.
     If the string is shorter than lengthOverride, the extra space is padded with
     zeroes.

     @param {string} str The string to write.
     @param {?boolean} endianness The endianness to use for the written string data.
     @param {?number} lengthOverride The number of characters to write.
     */
    writeUCS2String(str, endianness, lengthOverride) {
        if (lengthOverride == null) {
            lengthOverride = str.length;
        }
        for (let i = 0; i < str.length && i < lengthOverride; i++) {
            this.writeUint16(str.charCodeAt(i), endianness);
        }
        for (; i < lengthOverride; i++) {
            this.writeUint16(0);
        }
    }

    /**
     Read a string of desired length and encoding from the DataStream.

     @param {number} length The length of the string to read in bytes.
     @param {?string} encoding The encoding of the string data in the DataStream.
     Defaults to ASCII.
     @return {string} The read string.
     */
    readString(length, encoding) {
        if (encoding == null || encoding == "ASCII") {
            return DataStream.createStringFromArray(this.mapUint8Array(length == null ? this.byteLength - this.position : length));
        } else {
            return (new TextDecoder(encoding)).decode(this.mapUint8Array(length));
        }
    }

    /**
     Writes a string of desired length and encoding to the DataStream.

     @param {string} s The string to write.
     @param {?string} encoding The encoding for the written string data.
     Defaults to ASCII.
     @param {?number} length The number of characters to write.
     */
    writeString(s, encoding, length) {
        if (encoding == null || encoding == "ASCII") {
            if (length != null) {
                let i = 0;
                let len = Math.min(s.length, length);
                for (i = 0; i < len; i++) {
                    this.writeUint8(s.charCodeAt(i));
                }
                for (; i < length; i++) {
                    this.writeUint8(0);
                }
            } else {
                for (let i = 0; i < s.length; i++) {
                    this.writeUint8(s.charCodeAt(i));
                }
            }
        } else {
            this.writeUint8Array((new TextEncoder(encoding)).encode(s.substring(0, length)));
        }
    }

    /**
     Read null-terminated string of desired length from the DataStream. Truncates
     the returned string so that the null byte is not a part of it.

     @param {?number} length The length of the string to read.
     @return {string} The read string.
     */
    readCString(length) {
        let blen = this.byteLength - this.position;
        let u8 = new Uint8Array(this.#buffer, this.#byteOffset + this.position);
        let len = blen;
        if (length != null) {
            len = Math.min(length, blen);
        }
        for (let i = 0; i < len && u8[i] != 0; i++)
            ; // find first zero byte
        let s = DataStream.createStringFromArray(this.mapUint8Array(i));
        if (length != null) {
            this.position += len - i;
        } else if (i != blen) {
            this.position += 1; // trailing zero if not at end of buffer
        }
        return s;
    }

    /**
     Writes a null-terminated string to DataStream and zero-pads it to length
     bytes. If length is not given, writes the string followed by a zero.
     If string is longer than length, the written part of the string does not have
     a trailing zero.

     @param {string} s The string to write.
     @param {?number} length The number of characters to write.
     */
    writeCString(s, length) {
        if (length != null) {
            let i = 0;
            let len = Math.min(s.length, length);
            for (i = 0; i < len; i++) {
                this.writeUint8(s.charCodeAt(i));
            }
            for (; i < length; i++) {
                this.writeUint8(0);
            }
        } else {
            for (let i = 0; i < s.length; i++) {
                this.writeUint8(s.charCodeAt(i));
            }
            this.writeUint8(0);
        }
    }

    /**
     Reads an object of type t from the DataStream, passing struct as the thus-far
     read struct to possible callbacks that refer to it. Used by readStruct for
     reading in the values, so the type is one of the readStruct types.

     @param {Object} t Type of the object to read.
     @param {?Object} struct Struct to refer to when resolving length references
     and for calling callbacks.
     @return {?Object} Returns the object on successful read, null on unsuccessful.
     */
    readType(t, struct) {
        if (typeof t == "function") {
            return t(this, struct);
        } else if (typeof t == "object" && !(t instanceof Array)) {
            return t.get(this, struct);
        } else if (t instanceof Array && t.length != 3) {
            return this.readStruct(t, struct);
        }
        let v = null;
        let lengthOverride = null;
        let charset = "ASCII";
        let pos = this.position;
        let len;
        if (typeof t == 'string' && /:/.test(t)) {
            let tp = t.split(":");
            t = tp[0];
            len = tp[1];

            // allow length to be previously parsed variable
            // e.g. 'string:fieldLength', if `fieldLength` has
            // been parsed previously.
            if (struct[len] != null) {
                lengthOverride = parseInt(struct[len]);
            } else {
                // assume literal integer e.g., 'string:4'
                lengthOverride = parseInt(tp[1]);
            }
        }
        if (typeof t == 'string' && /,/.test(t)) {
            let tp = t.split(",");
            t = tp[0];
            charset = parseInt(tp[1]);
        }
        switch (t) {

            case 'uint8':
                v = this.readUint8();
                break;
            case 'int8':
                v = this.readInt8();
                break;

            case 'uint16':
                v = this.readUint16(this.endianness);
                break;
            case 'int16':
                v = this.readInt16(this.endianness);
                break;
            case 'uint32':
                v = this.readUint32(this.endianness);
                break;
            case 'int32':
                v = this.readInt32(this.endianness);
                break;
            case 'float32':
                v = this.readFloat32(this.endianness);
                break;
            case 'float64':
                v = this.readFloat64(this.endianness);
                break;

            case 'uint16be':
                v = this.readUint16(DataStream.BIG_ENDIAN);
                break;
            case 'int16be':
                v = this.readInt16(DataStream.BIG_ENDIAN);
                break;
            case 'uint32be':
                v = this.readUint32(DataStream.BIG_ENDIAN);
                break;
            case 'int32be':
                v = this.readInt32(DataStream.BIG_ENDIAN);
                break;
            case 'float32be':
                v = this.readFloat32(DataStream.BIG_ENDIAN);
                break;
            case 'float64be':
                v = this.readFloat64(DataStream.BIG_ENDIAN);
                break;

            case 'uint16le':
                v = this.readUint16(DataStream.LITTLE_ENDIAN);
                break;
            case 'int16le':
                v = this.readInt16(DataStream.LITTLE_ENDIAN);
                break;
            case 'uint32le':
                v = this.readUint32(DataStream.LITTLE_ENDIAN);
                break;
            case 'int32le':
                v = this.readInt32(DataStream.LITTLE_ENDIAN);
                break;
            case 'float32le':
                v = this.readFloat32(DataStream.LITTLE_ENDIAN);
                break;
            case 'float64le':
                v = this.readFloat64(DataStream.LITTLE_ENDIAN);
                break;

            case 'cstring':
                v = this.readCString(lengthOverride);
                break;

            case 'string':
                v = this.readString(lengthOverride, charset);
                break;

            case 'u16string':
                v = this.readUCS2String(lengthOverride, this.endianness);
                break;

            case 'u16stringle':
                v = this.readUCS2String(lengthOverride, DataStream.LITTLE_ENDIAN);
                break;

            case 'u16stringbe':
                v = this.readUCS2String(lengthOverride, DataStream.BIG_ENDIAN);
                break;

            default:
                if (t.length == 3) {
                    let ta = t[1];
                    let len = t[2];
                    let length = 0;
                    if (typeof len == 'function') {
                        length = len(struct, this, t);
                    } else if (typeof len == 'string' && struct[len] != null) {
                        length = parseInt(struct[len]);
                    } else {
                        length = parseInt(len);
                    }
                    if (typeof ta == "string") {
                        let tap = ta.replace(/(le|be)$/, '');
                        let endianness = null;
                        if (/le$/.test(ta)) {
                            endianness = DataStream.LITTLE_ENDIAN;
                        } else if (/be$/.test(ta)) {
                            endianness = DataStream.BIG_ENDIAN;
                        }
                        if (len == '*') {
                            length = null;
                        }
                        switch (tap) {
                            case 'uint8':
                                v = this.readUint8Array(length);
                                break;
                            case 'uint16':
                                v = this.readUint16Array(length, endianness);
                                break;
                            case 'uint32':
                                v = this.readUint32Array(length, endianness);
                                break;
                            case 'int8':
                                v = this.readInt8Array(length);
                                break;
                            case 'int16':
                                v = this.readInt16Array(length, endianness);
                                break;
                            case 'int32':
                                v = this.readInt32Array(length, endianness);
                                break;
                            case 'float32':
                                v = this.readFloat32Array(length, endianness);
                                break;
                            case 'float64':
                                v = this.readFloat64Array(length, endianness);
                                break;
                            case 'cstring':
                            case 'utf16string':
                            case 'string':
                                if (length == null) {
                                    v = [];
                                    while (!this.isEof()) {
                                        let u = this.readType(ta, struct);
                                        if (u == null)
                                            break;
                                        v.push(u);
                                    }
                                } else {
                                    v = new Array(length);
                                    for (let i = 0; i < length; i++) {
                                        v[i] = this.readType(ta, struct);
                                    }
                                }
                                break;
                        }
                    } else {
                        if (len == '*') {
                            v = [];
                            this.buffer;
                            while (true) {
                                let p = this.position;
                                try {
                                    let o = this.readType(ta, struct);
                                    if (o == null) {
                                        this.position = p;
                                        break;
                                    }
                                    v.push(o);
                                } catch (e) {
                                    this.position = p;
                                    break;
                                }
                            }
                        } else {
                            v = new Array(length);
                            for (let i = 0; i < length; i++) {
                                let u = this.readType(ta, struct);
                                if (u == null)
                                    return null;
                                v[i] = u;
                            }
                        }
                    }
                    break;
                }
        }
        if (lengthOverride != null) {
            this.position = pos + lengthOverride;
        }
        return v;
    }

    /**
     Writes a struct to the DataStream. Takes a structDefinition that gives the
     types and a struct object that gives the values. Refer to readStruct for the
     structure of structDefinition.

     @param {Object} structDefinition Type definition of the struct.
     @param {Object} struct The struct data object.
     */
    writeStruct(structDefinition, struct) {
        for (let i = 0; i < structDefinition.length; i += 2) {
            let t = structDefinition[i + 1];
            this.writeType(t, struct[structDefinition[i]], struct);
        }
    }

    /**
     Writes object v of type t to the DataStream.

     @param {Object} t Type of data to write.
     @param {Object} v Value of data to write.
     @param {Object} struct Struct to pass to write callback functions.
     */
    writeType(t, v, struct) {
        if (typeof t == "function") {
            return t(this, v);
        } else if (typeof t == "object" && !(t instanceof Array)) {
            return t.set(this, v, struct);
        }
        let lengthOverride = null;
        let charset = "ASCII";
        let pos = this.position;
        if (typeof (t) == 'string' && /:/.test(t)) {
            let tp = t.split(":");
            t = tp[0];
            lengthOverride = parseInt(tp[1]);
        }
        if (typeof t == 'string' && /,/.test(t)) {
            let tp = t.split(",");
            t = tp[0];
            charset = parseInt(tp[1]);
        }

        switch (t) {
            case 'uint8':
                this.writeUint8(v);
                break;
            case 'int8':
                this.writeInt8(v);
                break;

            case 'uint16':
                this.writeUint16(v, this.endianness);
                break;
            case 'int16':
                this.writeInt16(v, this.endianness);
                break;
            case 'uint32':
                this.writeUint32(v, this.endianness);
                break;
            case 'int32':
                this.writeInt32(v, this.endianness);
                break;
            case 'float32':
                this.writeFloat32(v, this.endianness);
                break;
            case 'float64':
                this.writeFloat64(v, this.endianness);
                break;

            case 'uint16be':
                this.writeUint16(v, DataStream.BIG_ENDIAN);
                break;
            case 'int16be':
                this.writeInt16(v, DataStream.BIG_ENDIAN);
                break;
            case 'uint32be':
                this.writeUint32(v, DataStream.BIG_ENDIAN);
                break;
            case 'int32be':
                this.writeInt32(v, DataStream.BIG_ENDIAN);
                break;
            case 'float32be':
                this.writeFloat32(v, DataStream.BIG_ENDIAN);
                break;
            case 'float64be':
                this.writeFloat64(v, DataStream.BIG_ENDIAN);
                break;

            case 'uint16le':
                this.writeUint16(v, DataStream.LITTLE_ENDIAN);
                break;
            case 'int16le':
                this.writeInt16(v, DataStream.LITTLE_ENDIAN);
                break;
            case 'uint32le':
                this.writeUint32(v, DataStream.LITTLE_ENDIAN);
                break;
            case 'int32le':
                this.writeInt32(v, DataStream.LITTLE_ENDIAN);
                break;
            case 'float32le':
                this.writeFloat32(v, DataStream.LITTLE_ENDIAN);
                break;
            case 'float64le':
                this.writeFloat64(v, DataStream.LITTLE_ENDIAN);
                break;

            case 'cstring':
                this.writeCString(v, lengthOverride);
                break;

            case 'string':
                this.writeString(v, charset, lengthOverride);
                break;

            case 'u16string':
                this.writeUCS2String(v, this.endianness, lengthOverride);
                break;

            case 'u16stringle':
                this.writeUCS2String(v, DataStream.LITTLE_ENDIAN, lengthOverride);
                break;

            case 'u16stringbe':
                this.writeUCS2String(v, DataStream.BIG_ENDIAN, lengthOverride);
                break;

            default:
                if (t.length === 3) {
                    let ta = t[1];
                    for (let i = 0; i < v.length; i++) {
                        this.writeType(ta, v[i]);
                    }
                    break;
                } else {
                    this.writeStruct(t, v);
                    break;
                }
        }
        if (lengthOverride) {
            this.position = pos;
            this.#realloc(lengthOverride);
            this.position = pos + lengthOverride;
        }
    }

}