// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files
// Browser-compatible implementation for writablestream
class StringBuf {
  constructor(options) {
    // Use Uint8Array instead of Buffer for browser compatibility
    this._buf = new Uint8Array((options && options.size) || 16384);
    this._encoding = (options && options.encoding) || 'utf8';

    // where in the buffer we are at
    this._inPos = 0;

    // for use by toBuffer()
    this._buffer = undefined;
  }

  get length() {
    return this._inPos;
  }

  get capacity() {
    return this._buf.length;
  }

  get buffer() {
    return this._buf;
  }

  toBuffer() {
    // return the current data as a single enclosing buffer
    if (!this._buffer) {
      this._buffer = new Uint8Array(this.length);
      this._buffer.set(this._buf.subarray(0, this.length));
    }
    return this._buffer;
  }

  toString() {
    // Convert Uint8Array to string using TextDecoder
    const decoder = new TextDecoder('utf-8');
    return decoder.decode(this._buf.subarray(0, this.length));
  }

  reset(position) {
    position = position || 0;
    this._buffer = undefined;
    this._inPos = position;
  }

  _grow(min) {
    let size = this._buf.length * 2;
    while (size < min) {
      size *= 2;
    }
    const buf = new Uint8Array(size);
    buf.set(this._buf);
    this._buf = buf;
  }

  addText(text) {
    this._buffer = undefined;

    // Use TextEncoder to convert string to Uint8Array
    const encoder = new TextEncoder('utf-8');
    const textBytes = encoder.encode(text);
    const textLength = textBytes.length;

    // Check if we need to grow the buffer
    if (this._inPos + textLength > this._buf.length) {
      this._grow(this._inPos + textLength);
    }

    // Copy the text bytes to the buffer
    this._buf.set(textBytes, this._inPos);
    this._inPos += textLength;
  }

  addStringBuf(inBuf) {
    if (inBuf.length) {
      this._buffer = undefined;

      // Check if we need to grow the buffer
      if (this.length + inBuf.length > this.capacity) {
        this._grow(this.length + inBuf.length);
      }

      // Copy the input buffer's data to this buffer
      // eslint-disable-next-line no-underscore-dangle
      this._buf.set(inBuf._buf.subarray(0, inBuf.length), this._inPos);
      this._inPos += inBuf.length;
    }
  }
}

module.exports = StringBuf;