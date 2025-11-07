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

    // ... 之前的事件监听器和其他方法保持不变 ...

    async downloadWordDocument() {
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
    async generateWordDocument() {
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
    separateVandVWProducts(products) {
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
    createProductTablesWithContinuousPagination(products, type) {
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
    groupProductsByCode(products) {
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
    createProductTable(products) {
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
    createTableCell(text, isHeader = false, width = "auto") {
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

    createProductCell(product, bgColor, width) {
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

    createColorCell(product, bgColor, width) {
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

    createSizesCell(product, bgColor, width) {
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

    // ... 之前的数据处理方法需要调整以匹配Word生成的需求
    processData(data) {
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