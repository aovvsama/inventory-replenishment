class InventoryReplenishmentGenerator {
    constructor() {
        this.sizeOrder = {
            'XS': 1, 'S': 2, 'M': 3, 'L': 4, 'XL': 5, 'XXL': 6,
            '28': 7, '29': 8, '30': 9, '31': 10, '32': 11, '33': 12,
            '34': 13, '36': 14, '38': 15, 'ONS': 16
        };
        this.processedData = null;
        this.originalData = null;
        
        console.log('🚀 初始化库存补货生成器...');
        this.initializeEventListeners();
    }

    // 事件监听器初始化方法
    initializeEventListeners = () => {
        console.log('🔄 初始化事件监听器...');
        
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');

        console.log('📋 元素状态:', {
            uploadArea: uploadArea ? '找到' : '未找到',
            fileInput: fileInput ? '找到' : '未找到'
        });

        if (!uploadArea || !fileInput) {
            console.error('❌ 错误: 必要的DOM元素未找到');
            return;
        }

        // 拖拽事件
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('dragover');
            console.log('🎯 拖拽经过上传区域');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
            console.log('🚪 拖拽离开上传区域');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            console.log('📥 文件放下, 文件数量:', files.length);
            
            if (files.length > 0) {
                this.handleFiles(files);
            }
        });

        // 文件选择事件
        fileInput.addEventListener('change', (e) => {
            console.log('📁 文件选择变化');
            const files = e.target.files;
            if (files.length > 0) {
                console.log('📄 选择的文件:', files[0].name);
                this.handleFiles(files);
            }
        });

        // 下载按钮事件
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            downloadBtn.addEventListener('click', () => {
                this.downloadWordDocument();
            });
        }

        console.log('✅ 事件监听器初始化完成');
    }

    // 处理文件的方法
    handleFiles = async (files) => {
        const file = files[0];
        console.log('🔄 开始处理文件:', file.name);
        
        if (!this.isExcelFile(file)) {
            alert('请上传Excel文件 (.xlsx 或 .xls 格式)');
            return;
        }

        try {
            this.showProgress();
            
            // 读取Excel文件
            const data = await this.readExcelFile(file);
            this.originalData = data;
            
            // 处理数据
            this.processedData = this.processData(data);
            
            // 显示结果
            this.showResult();
            
        } catch (error) {
            console.error('❌ 处理文件时出错:', error);
            alert('处理文件时出错: ' + error.message);
            this.resetApp();
        }
    }

    // 检查是否为Excel文件
    isExcelFile = (file) => {
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        const allowedExtensions = ['.xlsx', '.xls'];
        
        return allowedTypes.includes(file.type) || 
               allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    }

    // 显示进度
    showProgress = () => {
        const uploadArea = document.getElementById('uploadArea');
        const progressArea = document.getElementById('progressArea');
        
        if (uploadArea) uploadArea.style.display = 'none';
        if (progressArea) progressArea.style.display = 'block';
        
        this.updateProgress('正在读取Excel文件...', 30);
    }

    // 更新进度
    updateProgress = (text, percent) => {
        const progressText = document.getElementById('progressText');
        const progressFill = document.getElementById('progressFill');
        
        if (progressText) progressText.textContent = text;
        if (progressFill) progressFill.style.width = percent + '%';
    }

    // 显示结果
    showResult = () => {
        const progressArea = document.getElementById('progressArea');
        const resultArea = document.getElementById('resultArea');
        const resultText = document.getElementById('resultText');
        
        if (progressArea) progressArea.style.display = 'none';
        if (resultArea) resultArea.style.display = 'block';
        
        if (resultText && this.processedData) {
            const vProducts = this.processedData.filter(item => 
                item.code && item.code.startsWith('V') && !item.code.startsWith('VW')
            );
            const vwProducts = this.processedData.filter(item => 
                item.code && item.code.startsWith('VW')
            );
            
            resultText.innerHTML = `
                成功处理 ${this.processedData.length} 个产品组合<br>
                V系列: ${vProducts.length} 个<br>
                VW系列: ${vwProducts.length} 个<br>
                <small>点击下方按钮下载Word文档</small>
            `;
        }
    }

    // 重置应用
    resetApp = () => {
        const uploadArea = document.getElementById('uploadArea');
        const progressArea = document.getElementById('progressArea');
        const resultArea = document.getElementById('resultArea');
        const fileInput = document.getElementById('fileInput');
        
        if (uploadArea) uploadArea.style.display = 'block';
        if (progressArea) progressArea.style.display = 'none';
        if (resultArea) resultArea.style.display = 'none';
        if (fileInput) fileInput.value = '';
        
        this.processedData = null;
        this.originalData = null;
    }

    // 读取Excel文件
    readExcelFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    this.updateProgress('解析Excel数据...', 60);
                    
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    console.log('✅ Excel文件读取成功，行数:', jsonData.length);
                    resolve(jsonData);
                } catch (error) {
                    console.error('❌ 解析Excel文件失败:', error);
                    reject(new Error('解析Excel文件失败: ' + error.message));
                }
            };
            
            reader.onerror = () => {
                console.error('❌ 读取文件失败');
                reject(new Error('读取文件失败'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }

    // 处理数据
    processData = (data) => {
        console.log('🔄 开始处理数据...');
        this.updateProgress('处理库存数据...', 80);
        
        if (!data || data.length < 2) {
            throw new Error('Excel数据为空或格式不正确');
        }

        // 获取表头
        const headers = data[0];
        const codeIndex = headers.findIndex(h => h === '商品条码' || h === '条码');
        const colorIndex = headers.findIndex(h => h === 'Color' || h === '颜色');
        const sizeIndex = headers.findIndex(h => h === 'Size' || h === '尺码');
        const stockIndex = headers.findIndex(h => h === '总库存' || h === '库存');

        if (codeIndex === -1) {
            throw new Error('Excel文件中找不到"商品条码"列');
        }

        // 处理数据行
        const processed = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row && row[codeIndex]) {
                processed.push({
                    code: String(row[codeIndex]).trim(),
                    color: colorIndex !== -1 && row[colorIndex] ? String(row[colorIndex]).trim() : '',
                    size: sizeIndex !== -1 && row[sizeIndex] ? String(row[sizeIndex]).trim() : '',
                    stock: stockIndex !== -1 && row[stockIndex] ? Number(row[stockIndex]) || 0 : 0
                });
            }
        }

        console.log('✅ 数据处理完成，产品数量:', processed.length);
        return processed;
    }

    // 下载Word文档
    downloadWordDocument = async () => {
        try {
            this.updateProgress('生成Word文档...', 90);
            
            if (!this.processedData || this.processedData.length === 0) {
                throw new Error('没有可处理的数据');
            }

            const doc = await this.generateWordDocument();
            const blob = await docx.Packer.toBlob(doc);
            
            // 下载文件
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `库存补货清单_${new Date().toISOString().split('T')[0]}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            console.log('✅ Word文档下载完成');
            
        } catch (error) {
            console.error('❌ 生成Word文档时出错:', error);
            alert('生成Word文档时出错: ' + error.message);
        }
    }

    // Word文档生成核心方法
    generateWordDocument = async () => {
        if (!this.processedData || this.processedData.length === 0) {
            throw new Error('没有数据可生成文档');
        }

        console.log('📝 开始生成Word文档...');

        // 分离V和VW系列产品
        const separatedProducts = this.separateVandVWProducts(this.processedData);
        const vProducts = separatedProducts.vProducts;
        const vwProducts = separatedProducts.vwProducts;
        
        const docChildren = [];
        let currentPage = 1;

        // 添加标题页
        docChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "库存补货清单 - 第" + currentPage + "页",
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
                        text: `统计信息：共 ${this.processedData.length} 个产品组合（V系列: ${vProducts.length}个, VW系列: ${vwProducts.length}个）`,
                        size: 16,
                        font: "微软雅黑",
                        bold: true
                    })
                ],
                spacing: { after: 400 }
            })
        );

        // 处理V系列产品 - 连续分页
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

            const vTables = this.createProductTablesWithContinuousPagination(vProducts, "V");
            vTables.forEach((table, index) => {
                if (index > 0) {
                    currentPage++;
                    docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
                    docChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: "MENS audit list - 第" + currentPage + "页",
                                    bold: true,
                                    size: 18,
                                    font: "微软雅黑"
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 200 }
                        })
                    );
                }
                docChildren.push(table);
            });
        }

        // 处理VW系列产品 - 连续分页
        if (vwProducts.length > 0) {
            currentPage++;
            docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "WOMENS audit list - 第" + currentPage + "页",
                            bold: true,
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 200 }
                })
            );

            const vwTables = this.createProductTablesWithContinuousPagination(vwProducts, "VW");
            vwTables.forEach((table, index) => {
                if (index > 0) {
                    currentPage++;
                    docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
                    docChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: "WOMENS audit list - 第" + currentPage + "页",
                                    bold: true,
                                    size: 18,
                                    font: "微软雅黑"
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 200 }
                        })
                    );
                }
                docChildren.push(table);
            });
        }

        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        size: {
                            width: 11906, // A4横向
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

        console.log('✅ Word文档结构创建完成');
        return doc;
    }

    // 分离V和VW系列产品
    separateVandVWProducts = (products) => {
        const vProducts = [];
        const vwProducts = [];
        
        products.forEach(product => {
            if (product.code && product.code.startsWith('VW')) {
                vwProducts.push(product);
            } else if (product.code && product.code.startsWith('V')) {
                vProducts.push(product);
            }
        });
        
        // 按商品条码排序
        vProducts.sort((a, b) => (a.code || '').localeCompare(b.code || ''));
        vwProducts.sort((a, b) => (a.code || '').localeCompare(b.code || ''));
        
        return { vProducts, vwProducts };
    }

    // 创建连续分页的表格
    createProductTablesWithContinuousPagination = (products, type) => {
        const groupedProducts = this.groupProductsByCode(products);
        const tables = [];
        
        // 每页显示的产品数量（每页左右两栏各20个产品）
        const productsPerPage = 40;
        const totalPages = Math.ceil(groupedProducts.length / productsPerPage);

        for (let page = 0; page < totalPages; page++) {
            const startIdx = page * productsPerPage;
            const endIdx = startIdx + productsPerPage;
            const pageProducts = groupedProducts.slice(startIdx, endIdx);
            
            const table = this.createProductTable(pageProducts);
            tables.push(table);
        }

        return tables;
    }

    // 按商品条码分组
    groupProductsByCode = (products) => {
        const groups = {};
        
        products.forEach(product => {
            if (!product.code) return;
            
            if (!groups[product.code]) {
                groups[product.code] = {
                    code: product.code,
                    color: product.color,
                    sizes: []
                };
            }
            
            // 添加尺码信息，库存<2的标记为需要补货
            groups[product.code].sizes.push({
                size: product.size,
                stock: product.stock,
                needReplenish: (product.stock || 0) < 2
            });
        });
        
        // 对每个商品的尺码进行排序
        Object.values(groups).forEach(group => {
            group.sizes.sort((a, b) => {
                const orderA = this.sizeOrder[a.size] || 999;
                const orderB = this.sizeOrder[b.size] || 999;
                return orderA - orderB;
            });
        });
        
        return Object.values(groups);
    }

    // 创建产品表格
    createProductTable = (products) => {
        const productsPerColumn = Math.ceil(products.length / 2);
        const tableRows = [];

        // 表头行
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

        // 数据行
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

    // 创建表格单元格的辅助方法
    createTableCell = (text, isHeader = false, width = "auto") => {
        return new docx.TableCell({
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            bold: isHeader,
                            size: isHeader ? 20 : 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER
                })
            ],
            shading: isHeader ? {
                fill: "E8E8E8"
            } : undefined,
            width: {
                size: width === "auto" ? 0 : parseInt(width),
                type: width === "auto" ? docx.WidthType.AUTO : docx.WidthType.PERCENTAGE
            },
            verticalAlign: docx.VerticalAlign.CENTER
        });
    }

    createProductCell = (product, bgColor, width) => {
        if (!product) {
            return new docx.TableCell({
                children: [new docx.Paragraph({})],
                shading: { fill: bgColor },
                width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE }
            });
        }

        return new docx.TableCell({
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: product.code || "",
                            size: 18,
                            font: "微软雅黑",
                            bold: true
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER
                })
            ],
            shading: { fill: bgColor },
            width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.CENTER
        });
    }

    createColorCell = (product, bgColor, width) => {
        if (!product) {
            return new docx.TableCell({
                children: [new docx.Paragraph({})],
                shading: { fill: bgColor },
                width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE }
            });
        }

        return new docx.TableCell({
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: product.color || "",
                            size: 18,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER
                })
            ],
            shading: { fill: bgColor },
            width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.CENTER
        });
    }

    createSizesCell = (product, bgColor, width) => {
        if (!product) {
            return new docx.TableCell({
                children: [new docx.Paragraph({})],
                shading: { fill: bgColor },
                width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE }
            });
        }

        const sizeTexts = [];
        product.sizes.forEach((sizeInfo, index) => {
            if (index > 0) sizeTexts.push(" ");
            
            const sizeText = `${sizeInfo.size}(${sizeInfo.stock})`;
            
            if (sizeInfo.needReplenish) {
                // 库存<2的显示为红色
                sizeTexts.push(
                    new docx.TextRun({
                        text: sizeText,
                        color: "FF0000", // 红色
                        bold: true,
                        size: 18,
                        font: "微软雅黑"
                    })
                );
            } else {
                sizeTexts.push(
                    new docx.TextRun({
                        text: sizeText,
                        size: 18,
                        font: "微软雅黑"
                    })
                );
            }
        });

        return new docx.TableCell({
            children: [
                new docx.Paragraph({
                    children: sizeTexts,
                    alignment: docx.AlignmentType.LEFT
                })
            ],
            shading: { fill: bgColor },
            width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.CENTER
        });
    }
}

// 全局函数
function resetApp() {
    if (window.generator) {
        window.generator.resetApp();
    }
}

// 页面加载完成后初始化
console.log('📄 页面加载状态:', document.readyState);

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        console.log('✅ DOM内容加载完成，初始化应用...');
        window.generator = new InventoryReplenishmentGenerator();
    });
} else {
    console.log('✅ DOM已就绪，立即初始化应用...');
    window.generator = new InventoryReplenishmentGenerator();
}