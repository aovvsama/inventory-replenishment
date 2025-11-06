class InventoryReplenishmentGenerator {
    constructor() {
        this.sizeOrder = {
            'XS': 1, 'S': 2, 'M': 3, 'L': 4, 'XL': 5, 'XXL': 6,
            '28': 7, '29': 8, '30': 9, '31': 10, '32': 11, '33': 12,
            '34': 13, '36': 14, '38': 15, 'ONS': 16
        };
        this.processedData = null;
        this.originalData = null;
        
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFileSelect(e.target.files[0]);
            }
        });

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                this.handleFileSelect(e.dataTransfer.files[0]);
            }
        });
    }

    async handleFileSelect(file) {
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            alert('请选择Excel文件（.xlsx 或 .xls 格式）');
            return;
        }

        this.showProgress();
        this.updateProgress(10, '正在读取Excel文件...');

        try {
            const data = await this.readExcelFile(file);
            this.originalData = data;
            this.updateProgress(40, '正在处理数据...');
            
            const processedData = this.processData(data);
            this.updateProgress(70, '正在生成文档...');
            
            this.processedData = processedData;
            this.updateProgress(100, '处理完成！');
            
            this.showResult(processedData);
        } catch (error) {
            this.showError(error.message);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    resolve(jsonData);
                } catch (error) {
                    reject(new Error('读取Excel文件失败：' + error.message));
                }
            };
            
            reader.onerror = function() {
                reject(new Error('读取文件失败'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    processData(data) {
        if (!data || data.length === 0) {
            throw new Error('Excel文件中没有数据');
        }

        console.log('原始数据:', data);
        console.log('所有列名:', Object.keys(data[0]));

        const cleanedData = data.filter(row => {
            const hasColor = row['Color'] && row['Color'].toString().trim() !== '';
            const hasSize = row['Size'] && row['Size'].toString().trim() !== '';
            const hasStock = parseFloat(row['总库存']) > 0;
            
            return hasColor && hasSize && hasStock;
        });

        console.log('清理后数据:', cleanedData);

        if (cleanedData.length === 0) {
            throw new Error('清理后没有有效数据，请检查Excel文件内容');
        }

        const columnMap = {
            productCode: '商品条码',
            color: 'Color', 
            size: 'Size',
            stock: '总库存'
        };

        const availableColumns = Object.keys(cleanedData[0]);
        const missingColumns = [];
        
        for (const [key, columnName] of Object.entries(columnMap)) {
            if (!availableColumns.includes(columnName)) {
                missingColumns.push(columnName);
            }
        }

        if (missingColumns.length > 0) {
            throw new Error('缺少必需的列：' + missingColumns.join(', ') + '\n\n检测到的列名：' + availableColumns.join(', '));
        }

        console.log('列映射验证通过:', columnMap);
        console.log('清理后有效数据:', cleanedData.length, '行');

        const grouped = {};
        cleanedData.forEach(row => {
            const productCode = row[columnMap.productCode].toString().trim();
            const color = row[columnMap.color].toString().trim();
            const size = row[columnMap.size].toString().trim();
            const stock = parseFloat(row[columnMap.stock]) || 0;
            
            const key = productCode + '_' + color;
            if (!grouped[key]) {
                grouped[key] = {
                    productCode: productCode,
                    color: color,
                    sizes: new Map()
                };
            }
            grouped[key].sizes.set(size, stock);
        });

        const processedData = Object.values(grouped).map(item => ({
            productCode: item.productCode,
            color: item.color,
            sizes: this.sortSizesWithStock(Array.from(item.sizes.entries()))
        }));

        console.log('处理后产品数量:', processedData.length);
        console.log('处理后的数据:', processedData);

        if (processedData.length === 0) {
            throw new Error('没有找到有效的产品数据');
        }

        return processedData;
    }

    sortSizesWithStock(sizeStockPairs) {
        const getSizeKey = (size) => {
            const sizeStr = size.toUpperCase().trim();
            
            if (this.sizeOrder.hasOwnProperty(sizeStr)) {
                return this.sizeOrder[sizeStr];
            }
            
            if (!isNaN(sizeStr) && sizeStr !== '') {
                return parseInt(sizeStr) + 100;
            }
            
            return 999;
        };

        return sizeStockPairs.sort((a, b) => getSizeKey(a[0]) - getSizeKey(b[0]));
    }

    sortProductsByCode(products) {
        return products.sort((a, b) => {
            const codeA = a.productCode.toString();
            const codeB = b.productCode.toString();
            return codeA.localeCompare(codeB);
        });
    }

    separateVandVWProducts(data) {
        const vProducts = [];
        const vwProducts = [];

        data.forEach(item => {
            const productCode = item.productCode.toString();
            if (productCode.startsWith('VW')) {
                vwProducts.push(item);
            } else {
                vProducts.push(item);
            }
        });

        return {
            vProducts: this.sortProductsByCode(vProducts),
            vwProducts: this.sortProductsByCode(vwProducts)
        };
    }

    // 按阅读顺序重新组织产品（先左栏后右栏）
    organizeForReadingOrder(products) {
        const groupedProducts = this.groupProductsByCode(products);
        const result = [];
        const totalProducts = groupedProducts.length;
        const half = Math.ceil(totalProducts / 2);
        
        // 左栏：前一半产品
        for (let i = 0; i < half; i++) {
            result.push(groupedProducts[i]);
        }
        
        // 右栏：后一半产品
        for (let i = half; i < totalProducts; i++) {
            result.push(groupedProducts[i]);
        }
        
        return result;
    }

    groupProductsByCode(products) {
        const grouped = {};
        products.forEach(product => {
            if (!grouped[product.productCode]) {
                grouped[product.productCode] = [];
            }
            grouped[product.productCode].push(product);
        });

        const result = [];
        Object.keys(grouped).forEach(code => {
            const items = grouped[code];
            
            result.push({
                ...items[0],
                showCode: true
            });
            
            for (let i = 1; i < items.length; i++) {
                result.push({
                    ...items[i],
                    showCode: false
                });
            }
        });

        return result;
    }

    async generateWordDocument() {
        const separatedProducts = this.separateVandVWProducts(this.processedData);
        const vProducts = separatedProducts.vProducts;
        const vwProducts = separatedProducts.vwProducts;
        
        const docChildren = [];
        let pageNumber = 1;

        // 添加标题
        docChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "库存补货清单 - 第" + pageNumber + "页",
                        bold: true,
                        size: 24,
                        font: "微软雅黑"
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 200 }
            })
        );

        const timestamp = new Date().toLocaleString('zh-CN');
        docChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "生成时间: " + timestamp,
                        size: 16,
                        font: "微软雅黑"
                    })
                ],
                alignment: docx.AlignmentType.LEFT,
                spacing: { after: 200 }
            })
        );

        docChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "统计信息：共 " + this.processedData.length + " 个产品组合（V系列: " + vProducts.length + "个, VW系列: " + vwProducts.length + "个）",
                        size: 16,
                        font: "微软雅黑",
                        bold: true
                    })
                ],
                spacing: { after: 200 }
            })
        );

        // 添加V系列产品
        if (vProducts.length > 0) {
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "MENS audit list",
                            bold: true,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 200 }
                })
            );

            // 按阅读顺序组织V系列产品
            const readingOrderVProducts = this.organizeForReadingOrder(vProducts);
            const vTable = this.createProductTable(readingOrderVProducts);
            docChildren.push(vTable);
        }

        // 添加分页符（如果有VW系列产品）
        if (vwProducts.length > 0) {
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.PageBreak()
                    ]
                })
            );

            pageNumber++;

            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "WOMENS audit list - 第" + pageNumber + "页",
                            bold: true,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 200 }
                })
            );

            // 按阅读顺序组织VW系列产品
            const readingOrderVWProducts = this.organizeForReadingOrder(vwProducts);
            const vwTable = this.createProductTable(readingOrderVWProducts);
            docChildren.push(vwTable);
        }

        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        size: {
                            width: 11906,
                            height: 16838
                        },
                        margin: {
                            top: 500,
                            right: 500,
                            bottom: 500,
                            left: 500
                        }
                    },
                    pageOrientation: docx.PageOrientation.LANDSCAPE
                },
                children: docChildren
            }]
        });

        return doc;
    }

    createProductTable(products) {
        const productsPerColumn = Math.ceil(products.length / 2);
        const tableRows = [];

        const headerRow = new docx.TableRow({
            children: [
                this.createTableCell("商品条码", true, "15%"),
                this.createTableCell("颜色", true, "8%"),
                this.createTableCell("尺码", true, "32%"),
                this.createTableCell("商品条码", true, "15%"),
                this.createTableCell("颜色", true, "8%"),
                this.createTableCell("尺码", true, "32%")
            ],
            height: {
                value: 350,
                rule: docx.HeightRule.EXACT
            }
        });
        tableRows.push(headerRow);

        for (let i = 0; i < productsPerColumn; i++) {
            const leftProduct = products[i];
            const rightProduct = products[i + productsPerColumn];
            const bgColor = i % 2 === 0 ? "F8F8F8" : "FFFFFF";

            const row = new docx.TableRow({
                children: [
                    this.createProductCell(leftProduct, bgColor, "15%"),
                    this.createColorCell(leftProduct, bgColor, "8%"),
                    this.createSizesCell(leftProduct, bgColor, "32%"),
                    this.createProductCell(rightProduct, bgColor, "15%"),
                    this.createColorCell(rightProduct, bgColor, "8%"),
                    this.createSizesCell(rightProduct, bgColor, "32%")
                ],
                height: {
                    value: 350,
                    rule: docx.HeightRule.EXACT
                }
            });
            tableRows.push(row);
        }

        const table = new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            columnWidths: [15, 8, 32, 15, 8, 32],
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" },
                bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" },
                left: { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" },
                right: { style: docx.BorderStyle.SINGLE, size: 1, color: "000000" },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 0.5, color: "E0E0E0" },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 0.5, color: "E0E0E0" }
            },
            rows: tableRows
        });

        return table;
    }

    createTableCell(text, isHeader, width) {
        const widthValue = parseInt(width);
        return new docx.TableCell({
            width: { size: widthValue, type: docx.WidthType.PERCENTAGE },
            margins: { 
                top: 50,
                bottom: 50, 
                left: 50, 
                right: 50 
            },
            shading: isHeader ? { fill: "E8E8E8" } : undefined,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 14,
                            font: "微软雅黑",
                            bold: isHeader
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { 
                        line: 200,
                        before: 0,
                        after: 0
                    }
                })
            ]
        });
    }

    createProductCell(product, bgColor, width) {
        const text = product ? (product.showCode ? product.productCode : "") : "";
        const cell = this.createTableCell(text, false, width);
        cell.shading = { fill: bgColor };
        return cell;
    }

    createColorCell(product, bgColor, width) {
        const text = product ? product.color : "";
        const cell = this.createTableCell(text, false, width);
        cell.shading = { fill: bgColor };
        return cell;
    }

    createSizesCell(product, bgColor, width) {
        if (!product) {
            return this.createTableCell("", false, width);
        }

        const paragraphChildren = [];
        const sizesWithStock = product.sizes;
        
        sizesWithStock.forEach((sizeStock, index) => {
            const size = sizeStock[0];
            const stock = sizeStock[1];
            const separator = index < sizesWithStock.length - 1 ? " " : "";
            
            // 修正：库存 < 2 时标红（需要补货）
            if (stock < 2) {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: size,
                        color: "FF0000", // 红色
                        size: 14,
                        font: "微软雅黑"
                    })
                );
            } else {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: size,
                        color: "000000", // 黑色
                        size: 14,
                        font: "微软雅黑"
                    })
                );
            }
            
            if (separator) {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: separator,
                        color: "000000",
                        size: 14,
                        font: "微软雅黑"
                    })
                );
            }
        });

        const cell = new docx.TableCell({
            width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE },
            margins: { 
                top: 50,
                bottom: 50, 
                left: 50, 
                right: 50 
            },
            shading: { fill: bgColor },
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    children: paragraphChildren,
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { 
                        line: 200,
                        before: 0,
                        after: 0
                    }
                })
            ]
        });

        return cell;
    }

    async downloadWordDocument() {
        try {
            this.updateProgress(80, '正在生成Word文档...');
            
            const doc = await this.generateWordDocument();
            
            const blob = await docx.Packer.toBlob(doc);
            const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            saveAs(blob, '补货清单_' + timestamp + '.docx');
            
        } catch (error) {
            console.error('生成Word文档失败:', error);
            alert('生成Word文档失败: ' + error.message);
        }
    }

    showProgress() {
        document.getElementById('uploadArea').style.display = 'none';
        document.getElementById('progressArea').style.display = 'block';
        document.getElementById('resultArea').style.display = 'none';
    }

    updateProgress(percent, text) {
        document.getElementById('progressFill').style.width = percent + '%';
        document.getElementById('progressText').textContent = text;
    }

    showResult(data) {
        const separatedProducts = this.separateVandVWProducts(data);
        const vProducts = separatedProducts.vProducts;
        const vwProducts = separatedProducts.vwProducts;
        
        const resultText = '处理完成！共找到 ' + data.length + ' 个产品组合\n' +
                         'V系列: ' + vProducts.length + '个\n' +
                         'VW系列: ' + vwProducts.length + '个\n\n' +
                         '✅ 优化特性：\n' +
                         '• 库存<2的尺码显示为红色（需要补货）\n' +
                         '• 添加页码标注\n' +
                         '• 按阅读顺序分栏（先左栏后右栏）\n' +
                         '• 保持单页连续表格布局';
        
        document.getElementById('resultText').innerHTML = resultText.replace(/\n/g, '<br>');
        
        document.getElementById('progressArea').style.display = 'none';
        document.getElementById('resultArea').style.display = 'block';
        
        document.getElementById('downloadBtn').onclick = () => this.downloadWordDocument();
    }

    showError(message) {
        document.getElementById('progressArea').style.display = 'none';
        document.getElementById('uploadArea').style.display = 'block';
        alert('错误：' + message);
    }
}

function resetApp() {
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadArea').style.display = 'block';
    document.getElementById('progressArea').style.display = 'none';
    document.getElementById('resultArea').style.display = 'none';
}

let app;
document.addEventListener('DOMContentLoaded', function() {
    app = new InventoryReplenishmentGenerator();
});