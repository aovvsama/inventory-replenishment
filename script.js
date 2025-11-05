class InventoryReplenishmentGenerator {
    constructor() {
        this.sizeOrder = {
            'XS': 1, 'S': 2, 'M': 3, 'L': 4, 'XL': 5, 'XXL': 6,
            '28': 7, '29': 8, '30': 9, '31': 10, '32': 11, '33': 12,
            '34': 13, '36': 14, '38': 15, 'ONS': 16
        };
        this.processedData = null;
        
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
            this.updateProgress(40, '正在处理数据...');
            
            const processedData = this.processData(data);
            this.updateProgress(70, '正在生成Word文档...');
            
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
            
            reader.onerror = () => reject(new Error('读取文件失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    processData(data) {
        if (!data || data.length === 0) {
            throw new Error('Excel文件中没有数据');
        }

        console.log('原始数据:', data);
        console.log('所有列名:', Object.keys(data[0]));

        // 先清理数据，删除Color或Size为空的测试行
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

        // 直接使用列名
        const columnMap = {
            productCode: '商品条码',
            color: 'Color', 
            size: 'Size',
            stock: '总库存'
        };

        // 验证列是否存在
        const availableColumns = Object.keys(cleanedData[0]);
        const missingColumns = [];
        
        for (const [key, columnName] of Object.entries(columnMap)) {
            if (!availableColumns.includes(columnName)) {
                missingColumns.push(columnName);
            }
        }

        if (missingColumns.length > 0) {
            throw new Error(`缺少必需的列：${missingColumns.join(', ')}\n\n检测到的列名：${availableColumns.join(', ')}`);
        }

        console.log('列映射验证通过:', columnMap);
        console.log('清理后有效数据:', cleanedData.length, '行');

        // 按商品条码和颜色分组
        const grouped = {};
        cleanedData.forEach(row => {
            const productCode = row[columnMap.productCode].toString().trim();
            const color = row[columnMap.color].toString().trim();
            const size = row[columnMap.size].toString().trim();
            
            const key = `${productCode}_${color}`;
            if (!grouped[key]) {
                grouped[key] = {
                    productCode: productCode,
                    color: color,
                    sizes: new Set()
                };
            }
            grouped[key].sizes.add(size);
        });

        // 转换并排序尺寸
        const processedData = Object.values(grouped).map(item => ({
            productCode: item.productCode,
            color: item.color,
            sizes: this.sortSizes(Array.from(item.sizes))
        }));

        console.log(`处理后产品数量: ${processedData.length}`);
        console.log('处理后的数据:', processedData);

        if (processedData.length === 0) {
            throw new Error('没有找到有效的产品数据');
        }

        return processedData;
    }

    sortSizes(sizes) {
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

        return sizes.sort((a, b) => getSizeKey(a) - getSizeKey(b));
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

        return { vProducts, vwProducts };
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

    formatSizesText(sizes, maxCharsPerLine = 25) {
        if (!sizes || sizes.length === 0) return "";
        
        const singleLine = sizes.join(" ");
        if (singleLine.length <= maxCharsPerLine) {
            return singleLine;
        }

        const lines = [];
        let currentLine = "";

        sizes.forEach(size => {
            if (!currentLine) {
                currentLine = size;
            } else {
                if (currentLine.length + size.length + 1 <= maxCharsPerLine) {
                    currentLine += " " + size;
                } else {
                    lines.push(currentLine);
                    currentLine = size;
                }
            }
        });

        if (currentLine) {
            lines.push(currentLine);
        }

        return lines.join("\n");
    }

    async generateWordDocument() {
        const { vProducts, vwProducts } = this.separateVandVWProducts(this.processedData);
        
        // 创建Word文档
        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 800,    // 约1.4cm
                            right: 800,  // 约1.4cm
                            bottom: 800, // 约1.4cm
                            left: 800,   // 约1.4cm
                        }
                    }
                },
                children: []
            }]
        });

        // 添加标题
        const title = new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "库存补货清单",
                    bold: true,
                    size: 28,
                    font: "微软雅黑"
                })
            ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 400 }
        });
        doc.addSection({
            properties: {},
            children: [title]
        });

        // 添加时间戳
        const timestamp = new Date().toLocaleString('zh-CN');
        const timeParagraph = new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `生成时间: ${timestamp}`,
                    size: 20,
                    font: "微软雅黑"
                })
            ],
            alignment: docx.AlignmentType.LEFT,
            spacing: { after: 600 }
        });
        doc.addParagraph(timeParagraph);

        // 添加统计信息
        const statsParagraph = new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `统计信息：共 ${this.processedData.length} 个产品组合（V系列: ${vProducts.length}个, VW系列: ${vwProducts.length}个）`,
                    size: 20,
                    font: "微软雅黑",
                    bold: true
                })
            ],
            spacing: { after: 400 }
        });
        doc.addParagraph(statsParagraph);

        // 添加V系列产品
        if (vProducts.length > 0) {
            this.addProductSection(doc, "MENS audit list", vProducts);
        }

        // 添加分页（如果有VW系列产品）
        if (vwProducts.length > 0) {
            doc.addSection({
                properties: {
                    page: {
                        margin: {
                            top: 800,
                            right: 800,
                            bottom: 800,
                            left: 800,
                        }
                    }
                },
                children: [
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "WOMENS audit list",
                                bold: true,
                                size: 24,
                                font: "微软雅黑"
                            })
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 400 }
                    })
                ]
            });

            this.addProductSection(doc, "", vwProducts, false);
        }

        return doc;
    }

    addProductSection(doc, title, products, addTitle = true) {
        if (addTitle && title) {
            const titleParagraph = new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: title,
                        bold: true,
                        size: 24,
                        font: "微软雅黑"
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 400 }
            });
            doc.addParagraph(titleParagraph);
        }

        const groupedProducts = this.groupProductsByCode(products);
        const productsPerColumn = Math.ceil(groupedProducts.length / 2);

        // 创建表格
        const table = new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE,
            },
            borders: {
                top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                insideHorizontal: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
                insideVertical: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
            rows: []
        });

        // 添加产品数据到表格
        for (let i = 0; i < productsPerColumn; i++) {
            const leftProduct = groupedProducts[i];
            const rightProduct = groupedProducts[i + productsPerColumn];

            const row = new docx.TableRow({
                children: []
            });

            // 左侧产品
            row.addCell(this.createProductCell(leftProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));
            row.addCell(this.createColorCell(leftProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));
            row.addCell(this.createSizesCell(leftProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));

            // 右侧产品
            row.addCell(this.createProductCell(rightProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));
            row.addCell(this.createColorCell(rightProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));
            row.addCell(this.createSizesCell(rightProduct, i % 2 === 0 ? "F5F5F5" : "FFFFFF"));

            table.addRow(row);
        }

        doc.addParagraph(new docx.Paragraph({
            children: [table],
            spacing: { after: 200 }
        }));
    }

    createProductCell(product, bgColor) {
        const text = product ? (product.showCode ? product.productCode : "") : "";
        return new docx.TableCell({
            width: { size: 15, type: docx.WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            shading: { fill: bgColor },
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT
                })
            ]
        });
    }

    createColorCell(product, bgColor) {
        const text = product ? product.color : "";
        return new docx.TableCell({
            width: { size: 12, type: docx.WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            shading: { fill: bgColor },
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT
                })
            ]
        });
    }

    createSizesCell(product, bgColor) {
        const text = product ? this.formatSizesText(product.sizes, 20) : "";
        return new docx.TableCell({
            width: { size: 23, type: docx.WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            shading: { fill: bgColor },
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT
                })
            ]
        });
    }

    async downloadWordDocument() {
        try {
            this.updateProgress(80, '正在生成Word文档...');
            
            const doc = await this.generateWordDocument();
            
            // 生成文档并下载
            const blob = await docx.Packer.toBlob(doc);
            const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            saveAs(blob, `补货清单_${timestamp}.docx`);
            
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
        const { vProducts, vwProducts } = this.separateVandVWProducts(data);
        const resultText = `处理完成！共找到 ${data.length} 个产品组合
V系列: ${vProducts.length}个
VW系列: ${vwProducts.length}个`;
        
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
document.addEventListener('DOMContentLoaded', () => {
    app = new InventoryReplenishmentGenerator();
});