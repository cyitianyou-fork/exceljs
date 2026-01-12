const zipjs = require('@zip.js/zip.js');

// 环境检测
const isBrowser = typeof window !== 'undefined' && typeof window.document !== 'undefined';
const isNode = typeof process !== 'undefined' && process.versions != null && process.versions.node != null;

// 兼容性处理
let TextEncoder;
let fetch;
let WritableStream;
let ReadableStream;
let atob;

if (isBrowser) {
    // 浏览器环境
    TextEncoder = window.TextEncoder;
    fetch = window.fetch;
    WritableStream = window.WritableStream;
    ReadableStream = window.ReadableStream;
    atob = window.atob;
} else if (isNode) {
    // Node.js 环境
    // 只使用 Node.js 内置模块，避免使用外部依赖
    try {
        const util = require('util');
        TextEncoder = util.TextEncoder;
    } catch (e) {
        // 忽略错误
    }
    try {
        const buffer = require('buffer');
        atob = (str) => buffer.Buffer.from(str, 'base64').toString('binary');
    } catch (e) {
        // 忽略错误
    }
    // 注意：在 Node.js 环境中，fetch、WritableStream 和 ReadableStream 可能不可用
    // 但这些在浏览器环境中是必需的，Node.js 环境通常使用其他方式处理流
}

// 确保必要的 API 存在
if (!TextEncoder) {
    // 降级实现
    TextEncoder = function() {};
    TextEncoder.prototype.encode = function(str) {
        const bytes = [];
        for (let i = 0; i < str.length; i++) {
            bytes.push(str.charCodeAt(i));
        }
        return new Uint8Array(bytes);
    };
}

if (!atob) {
    // 降级实现
    atob = function(str) {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
        let result = '';
        let i = 0;
        str = str.replace(/[^A-Za-z0-9+/=]/g, '');
        while (i < str.length) {
            const a = chars.indexOf(str.charAt(i++));
            const b = chars.indexOf(str.charAt(i++));
            const c = chars.indexOf(str.charAt(i++));
            const d = chars.indexOf(str.charAt(i++));
            const ac = (a << 2) | (b >> 4);
            const bd = ((b & 15) << 4) | (c >> 2);
            const ce = ((c & 3) << 6) | d;
            result += String.fromCharCode(ac);
            if (c !== 64) result += String.fromCharCode(bd);
            if (d !== 64) result += String.fromCharCode(ce);
        }
        return result;
    };
}

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
class ZipWriter {
    constructor(options) {
        options = options || {};
        zipjs.configure({
            useWebWorkers: options.useWebWorkers || false,
        });
        this.zipWriterStream = new zipjs.ZipWriterStream(Object.assign({
            zip64: false,
            compressionMethod: 8,
        }, options.zipOptions || {}));
        if (options.stream) {
            this.stream = options.stream;
        }
        // Check if stream is a standard WritableStream
        if (this.stream && this.stream instanceof WritableStream) {
            // Use pipeTo for standard WritableStream
            this.zipWriterStream.readable.pipeTo(this.stream);
        } else if (this.stream && typeof this.stream.write === 'function') {
            // Use custom write method for non-standard streams (like FileStream)
            this._pipeToCustomStream();
        }
        // 获取并发下载限制，确保是有效的数字
        this.maxConcurrentDownloads = Math.max(1, parseInt(options.maxConcurrentDownloads, 10) || 10);
        // 初始化下载队列
        this.downloadQueue = [];
        this.activeDownloads = 0;
        this.onImageDownload = options.onImageDownload || ((imgUrl, isSuccess) => {});
        this.promises = [];
    }

    // 处理下载队列
    _processDownloadQueue() {
        while (this.activeDownloads < this.maxConcurrentDownloads && this.downloadQueue.length > 0) {
            const task = this.downloadQueue.shift();
            this.activeDownloads++;
            task()
                .finally(() => {
                    this.activeDownloads--;
                    this._processDownloadQueue();
                });
        }
    }

    // 添加下载任务到队列
    _addDownloadTask(task) {
        return new Promise((resolve, reject) => {
            this.downloadQueue.push(async() => {
                try {
                    const result = await task();
                    resolve(result);
                    return result;
                } catch (error) {
                    reject(error);
                    throw error;
                }
            });
            this._processDownloadQueue();
        });
    }

    async _pipeToCustomStream() {
        try {
            const reader = this.zipWriterStream.readable.getReader();
            while (true) {
                const { done, value } = await reader.read();
                if (done) {
                    if (typeof this.stream.end === 'function') {
                        await this.stream.end();
                    }
                    break;
                }
                await this.stream.write(value);
            }
        } catch (error) {
            console.error('Error piping to custom stream:', error);
            throw error;
        }
    }
    _openStream(path) {
        // Normalize file path - remove leading slash to ensure consistent paths
        let normalizedPath = path.replace(/^\//, '');

        this.createFolder(normalizedPath);

        const writable = this.zipWriterStream.writable(normalizedPath);
        const writer = writable.getWriter();

        // Save original write method to avoid recursion
        const originalWrite = writer.write;

        // Add write and end methods to mimic the old stream interface
        writer.write = (data) => {
            let promise;
            if (typeof data === 'string') {
                const encoder = new TextEncoder('utf-8');
                const uint8Array = encoder.encode(data);
                promise = originalWrite.call(writer, uint8Array);
            } else if (data instanceof Uint8Array) {
                promise = originalWrite.call(writer, data);
            } else if (data.reset && data.addText) {
                // Handle StringBuf objects
                const encoder = new TextEncoder('utf-8');
                let uint8Array;
                if (typeof data.toString === 'function') {
                    uint8Array = encoder.encode(data.toString());
                } else if (data.buffer instanceof Uint8Array) {
                    // Handle our browser-compatible StringBuf
                    uint8Array = data.buffer.subarray(0, data.length);
                } else {
                    // Fallback to string conversion
                    uint8Array = encoder.encode(String(data));
                }
                promise = originalWrite.call(writer, uint8Array);
            }
            // Add the promise to the array so we can wait for it in finalize
            if (promise) {
                this.promises.push(promise);
            }
            return promise;
        };
        writer.end = () => {
            const promise = writer.close();
            // Add the promise to the array so we can wait for it in finalize
            if (promise) {
                this.promises.push(promise);
            }
            return promise;
        };
        return writer;
    }
    async append(data, options) {
        // Normalize file path - remove leading slash to ensure consistent paths
        let normalizedName = options.name.replace(/^\//, '');

        await this.createFolder(normalizedName);

        let closePromise;
        if (options.hasOwnProperty('base64') && options.base64) {
            // For base64 data, decode it first
            const binaryString = atob(data);
            const len = binaryString.length;
            const uint8Array = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                uint8Array[i] = binaryString.charCodeAt(i);
            }
            const writable = this.zipWriterStream.writable(normalizedName);
            const writer = writable.getWriter();
            const writePromise = writer.write(uint8Array);
            closePromise = writePromise.then(() => writer.close());
        } else if (options.hasOwnProperty('imgUrl') && options.imgUrl) {
            // For image URLs, fetch the image data first using Promise chain with concurrency control
            closePromise = this._addDownloadTask(async() => {
                try {
                    const response = await fetch(data);
                    let uint8Array, isSuccess = true;
                    if (!response.ok) {
                        uint8Array = new Uint8Array(0);
                        isSuccess = false;
                    } else if (response.body instanceof ReadableStream) {
                        // 使用 response.bytes() 方法获取 Uint8Array
                        uint8Array = await response.bytes();
                    } else {
                        // 降级方案：使用 arrayBuffer 并转换为 Uint8Array
                        const arrayBuffer = await response.arrayBuffer();
                        uint8Array = new Uint8Array(arrayBuffer);
                    }
                    // 通知外部处理函数，如更新进度条
                    if (this.onImageDownload) {
                        this.onImageDownload(data, isSuccess);
                    }
                    const writable = this.zipWriterStream.writable(normalizedName);
                    const writer = writable.getWriter();
                    const writePromise = writer.write(uint8Array);
                    return writePromise.then(() => writer.close());
                } catch (error) {
                    // 下载失败时返回空数据，避免整个导出过程失败
                    const writable = this.zipWriterStream.writable(normalizedName);
                    const writer = writable.getWriter();
                    const writePromise = writer.write(new Uint8Array(0));
                    return writePromise.then(() => writer.close());
                }
            });
        } else {
            const writable = this.zipWriterStream.writable(normalizedName);
            const writer = writable.getWriter();
            let writePromise;
            if (typeof data === 'string') {
                // Use TextEncoder in browser directly
                const encoder = new TextEncoder('utf-8');
                const uint8Array = encoder.encode(data);
                writePromise = writer.write(uint8Array);
            } else if (data instanceof ArrayBuffer) {
                writePromise = writer.write(new Uint8Array(data));
            } else if (data instanceof Uint8Array) {
                writePromise = writer.write(data);
            }
            closePromise = writePromise ? writePromise.then(() => writer.close()) : writer.close();
        }
        this.promises.push(closePromise);
        await closePromise;
    }

    async createFolder(filePath) {
        let zipWriter = this.zipWriterStream.zipWriter;
        let tempPath = "";
        let parts = filePath.split("/");
        if (parts.length > 0) {
            parts.pop();
        }
        const folderPromises = [];
        for (let part of parts) {
            tempPath += part + "/";
            if (zipWriter.filenames.has(tempPath)) {
                continue;
            }
            // 使用 writable 方法创建目录，符合 @zip.js/zip.js 库的设计
            const writable = this.zipWriterStream.writable(tempPath);
            const writer = writable.getWriter();
            // 写入空内容作为目录条目
            const writePromise = writer.write(new Uint8Array(0));
            const closePromise = writePromise ? writePromise.then(() => writer.close()) : writer.close();
            // 收集 Promise 以确保目录创建完成
            this.promises.push(closePromise);
            folderPromises.push(closePromise);
        }
        // 一次性等待所有目录创建完成
        await Promise.all(folderPromises);
    }
    async finalize() {
        // Wait for all promises to resolve before closing the zip
        await Promise.all(this.promises);
        // Close the zip writer stream to finalize the zip file
        await this.zipWriterStream.close();
    }
}

// =============================================================================

module.exports = {
    ZipWriter,
};