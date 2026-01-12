const RelType = require('../xlsx/rel-type');
const StylesXform = require('../xlsx/xform/style/styles-xform');
const SharedStrings = require('../utils/shared-strings');
const DefinedNames = require('../doc/defined-names');

const CoreXform = require('../xlsx/xform/core/core-xform');
const RelationshipsXform = require('../xlsx/xform/core/relationships-xform');
const ContentTypesXform = require('../xlsx/xform/core/content-types-xform');
const AppXform = require('../xlsx/xform/core/app-xform');
const WorkbookXform = require('../xlsx/xform/book/workbook-xform');
const SharedStringsXform = require('../xlsx/xform/strings/shared-strings-xform');

const WorksheetWriter = require('./worksheet-writer');

const { ZipWriter } = require('./zip-writer');

const theme1Xml = require('../xlsx/xml/theme1');

class WorkbookWriter {
    constructor(options) {
        options = options || {};

        this.created = options.created || new Date();
        this.modified = options.modified || this.created;
        this.creator = options.creator || 'ExcelJS';
        this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
        this.lastPrinted = options.lastPrinted;

        // using shared strings creates a smaller xlsx file but may use more memory
        this.useSharedStrings = options.useSharedStrings || false;
        this.sharedStrings = new SharedStrings();

        // style manager
        this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock(true);

        // defined names
        this._definedNames = new DefinedNames();

        this._worksheets = [];
        this.views = [];

        this.media = [];
        this.drawings = [];
        this.commentRefs = [];
        this.stream = options.stream;
        this.zip = new ZipWriter(options);
        this.imageDelayDownload = !!options.imageDelayDownload;
        // Initialize promise with empty array to maintain compatibility
        this.promise = Promise.resolve([]);
    }

    get definedNames() {
        return this._definedNames;
    }

    _openStream(path) {
        const stream = this.zip._openStream(path);
        return stream;
    }

    _commitWorksheets() {
        // commit all uncommitted worksheets with final commit
        this._worksheets.forEach(worksheet => {
            if (!worksheet.committed) {
                worksheet.finalCommit();
            }
        });
        return Promise.resolve();
    }

    async commit() {
        // commit all worksheets, then add suplimentary files
        await this.promise;
        await this.addMedia();
        await this._commitWorksheets();
        await this.addDrawings();
        await Promise.all([
            this.addContentTypes(),
            this.addApp(),
            this.addCore(),
            this.addSharedStrings(),
            this.addStyles(),
            this.addWorkbookRels(),
            this.addThemes(),
            this.addOfficeRels(),
        ]);
        await this.addWorkbook();
        return this._finalize();
    }

    get nextId() {
        // find the next unique spot to add worksheet
        let i;
        for (i = 1; i < this._worksheets.length; i++) {
            if (!this._worksheets[i]) {
                return i;
            }
        }
        return this._worksheets.length || 1;
    }

    addImage(image) {
        const id = this.media.length;
        const medium = Object.assign({}, image, { type: 'image', name: `image${id}.${image.extension}` });
        this.media.push(medium);
        if (medium.imgUrl && !this.imageDelayDownload) {
            // 如果是图片 URL，需要特殊处理
            const filename = `xl/media/${medium.name}`;
            this.zip.append(medium.imgUrl, { name: filename, imgUrl: true });
        }
        return id;
    }

    getImage(id) {
        return this.media[id];
    }

    addWorksheet(name, options) {
        // it's possible to add a worksheet with different than default
        // shared string handling
        // in fact, it's even possible to switch it mid-sheet
        options = options || {};
        const useSharedStrings =
            options.useSharedStrings !== undefined ? options.useSharedStrings : this.useSharedStrings;

        if (options.tabColor) {
            // eslint-disable-next-line no-console
            console.trace('tabColor option has moved to { properties: tabColor: {...} }');
            options.properties = Object.assign({
                    tabColor: options.tabColor,
                },
                options.properties
            );
        }

        const id = this.nextId;
        name = name || `sheet${id}`;

        const worksheet = new WorksheetWriter({
            id,
            name,
            workbook: this,
            useSharedStrings,
            properties: options.properties,
            state: options.state,
            pageSetup: options.pageSetup,
            views: options.views,
            autoFilter: options.autoFilter,
            headerFooter: options.headerFooter,
        });

        this._worksheets[id] = worksheet;
        return worksheet;
    }

    getWorksheet(id) {
        if (id === undefined) {
            return this._worksheets.find(() => true);
        }
        if (typeof id === 'number') {
            return this._worksheets[id];
        }
        if (typeof id === 'string') {
            return this._worksheets.find(worksheet => worksheet && worksheet.name === id);
        }
        return undefined;
    }

    async addStyles() {
        await this.zip.append(this.styles.xml, { name: 'xl/styles.xml' });
    }

    async addThemes() {
        await this.zip.append(theme1Xml, { name: 'xl/theme/theme1.xml' });
    }

    async addOfficeRels() {
        const xform = new RelationshipsXform();
        const xml = xform.toXml([
            { Id: 'rId1', Type: RelType.OfficeDocument, Target: 'xl/workbook.xml' },
            { Id: 'rId2', Type: RelType.CoreProperties, Target: 'docProps/core.xml' },
            { Id: 'rId3', Type: RelType.ExtenderProperties, Target: 'docProps/app.xml' },
        ]);
        await this.zip.append(xml, { name: '/_rels/.rels' });
    }

    async addContentTypes() {
        const model = {
            worksheets: this._worksheets.filter(Boolean),
            sharedStrings: this.sharedStrings,
            commentRefs: this.commentRefs,
            media: this.media,
            drawings: this.drawings,
        };
        const xform = new ContentTypesXform();
        const xml = xform.toXml(model);
        await this.zip.append(xml, { name: '[Content_Types].xml' });
    }

    async addMedia() {

        await Promise.all(
            this.media.map(async medium => {
                if (medium.type === 'image') {
                    const filename = `xl/media/${medium.name}`;
                    if (medium.imgUrl) {
                        if (!this.imageDelayDownload) {
                            // 图片 URL 已在 addImage 中处理
                            return Promise.resolve();
                        } else {
                            await this.zip.append(medium.imgUrl, { name: filename, imgUrl: true });
                        }
                    } else if (medium.buffer) {
                        await this.zip.append(medium.buffer, { name: filename });
                    } else if (medium.base64) {
                        const dataimg64 = medium.base64;
                        const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
                        await this.zip.append(content, { name: filename, base64: true });
                    }
                } else {
                    throw new Error('Unsupported media');
                }
                return Promise.resolve();
            })
        );
    }
    async addDrawings() {
        if (this.drawings.length === 0) {
            return Promise.resolve();
        }

        const DrawingXform = require('../xlsx/xform/drawing/drawing-xform');
        const drawingXform = new DrawingXform();

        await Promise.all(
            this.drawings.map(drawing => {
                // Prepare the drawing before rendering
                drawingXform.prepare(drawing);

                // Generate drawing XML
                const xml = drawingXform.toXml(drawing);
                const drawingPath = `xl/drawings/${drawing.name}.xml`;

                // Generate drawing relationships XML if there are relationships
                const promises = [this.zip.append(xml, { name: drawingPath })];

                if (drawing.rels && drawing.rels.length > 0) {
                    const relationshipsXform = new RelationshipsXform();
                    const relsXml = relationshipsXform.toXml(drawing.rels);
                    const relsPath = `xl/drawings/_rels/${drawing.name}.xml.rels`;
                    promises.push(this.zip.append(relsXml, { name: relsPath }));
                }

                return Promise.all(promises);
            })
        );
        return Promise.resolve();
    }

    async addApp() {
        const model = {
            worksheets: this._worksheets.filter(Boolean),
        };
        const xform = new AppXform();
        const xml = xform.toXml(model);
        await this.zip.append(xml, { name: 'docProps/app.xml' });
    }

    async addCore() {
        const coreXform = new CoreXform();
        const xml = coreXform.toXml(this);
        await this.zip.append(xml, { name: 'docProps/core.xml' });
    }

    async addSharedStrings() {
        if (this.sharedStrings.count) {
            const sharedStringsXform = new SharedStringsXform();
            const xml = sharedStringsXform.toXml(this.sharedStrings);
            await this.zip.append(xml, { name: '/xl/sharedStrings.xml' });
        }
    }

    async addWorkbookRels() {
        let count = 1;
        const relationships = [
            { Id: `rId${count++}`, Type: RelType.Styles, Target: 'styles.xml' },
            { Id: `rId${count++}`, Type: RelType.Theme, Target: 'theme/theme1.xml' },
        ];
        if (this.sharedStrings.count) {
            relationships.push({
                Id: `rId${count++}`,
                Type: RelType.SharedStrings,
                Target: 'sharedStrings.xml',
            });
        }
        this._worksheets.forEach(worksheet => {
            if (worksheet) {
                worksheet.rId = `rId${count++}`;
                relationships.push({
                    Id: worksheet.rId,
                    Type: RelType.Worksheet,
                    Target: `worksheets/sheet${worksheet.id}.xml`,
                });
            }
        });
        const xform = new RelationshipsXform();
        const xml = xform.toXml(relationships);
        await this.zip.append(xml, { name: '/xl/_rels/workbook.xml.rels' });
    }

    async addWorkbook() {
        const model = {
            worksheets: this._worksheets.filter(Boolean),
            definedNames: this._definedNames.model,
            views: this.views,
            properties: {},
            calcProperties: {},
        };

        const xform = new WorkbookXform();
        xform.prepare(model);
        await this.zip.append(xform.toXml(model), { name: '/xl/workbook.xml' });
    }

    async _finalize() {
        await this.zip.finalize();
        return this;
    }
}

module.exports = WorkbookWriter;