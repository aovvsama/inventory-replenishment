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

    // ... 其他方法保持不变 ...

    async generateWordDocument() {
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
                        text: "统计信息：共 " + this.processedData.length + " 个产品组合（V系列: " + vProducts.length + "个, VW系列: " + vwProducts.length + "个）",
                        size: 16,
                        font: "微软雅黑",
                        bold: true
                    })
                ],
                spacing: { after: 400 }
            })
        );

        // 处理V系列产品 - 优化分页，每页显示更多产品
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

            const vTables = this.createProductTablesWithOptimizedPagination(vProducts, "V");
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

        // 处理VW系列产品 - 优化分页
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

            const vwTables = this.createProductTablesWithOptimizedPagination(vwProducts, "VW");
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

    // 优化分页：每页显示更多产品
    createProductTablesWithOptimizedPagination(products, type) {
        const groupedProducts = this.groupProductsByCode(products);
        const tables = [];
        
        // 优化：大幅增加每页显示的产品数量
        // 根据产品数量动态调整每页显示数量
        let productsPerPage;
        if (groupedProducts.length <= 60) {
            productsPerPage = 60; // 少量产品时，一页显示完
        } else if (groupedProducts.length <= 120) {
            productsPerPage = 80; // 中等数量产品
        } else {
            productsPerPage = 100; // 大量产品时，每页显示更多
        }
        
        const totalPages = Math.ceil(groupedProducts.length / productsPerPage);

        for (let page = 0; page < totalPages; page++) {
            const startIdx = page * productsPerPage;
            const endIdx = startIdx + productsPerPage;
            const pageProducts = groupedProducts.slice(startIdx, endIdx);
            
            const table = this.createProductTable(pageProducts);
            tables.push(table);
        }

        console.log(`📄 ${type}系列：共 ${groupedProducts.length} 个产品，分 ${totalPages} 页显示，每页 ${productsPerPage} 个产品`);
        return tables;
    }

    // 优化表格创建：减小行高，增加每页行数
    createProductTable(products) {
        const productsPerColumn = Math.ceil(products.length / 2);

        const tableRows = [];

        // 优化：减小表头行高
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
                value: 280, // 减小行高
                rule: docx.HeightRule.EXACT
            }
        });
        tableRows.push(headerRow);

        for (let i = 0; i < productsPerColumn; i++) {
            const leftProduct = products[i];
            const rightProduct = products[i + productsPerColumn];
            const bgColor = i % 2 === 0 ? "F8F8F8" : "FFFFFF";

            // 优化：减小数据行高
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
                    value: 280, // 减小行高
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

    // 优化单元格：减小字体和边距
    createTableCell(text, isHeader, width) {
        const widthValue = parseInt(width);
        return new docx.TableCell({
            width: { size: widthValue, type: docx.WidthType.PERCENTAGE },
            margins: { 
                top: 30,    // 减小边距
                bottom: 30, 
                left: 30, 
                right: 30 
            },
            shading: isHeader ? { fill: "E8E8E8" } : undefined,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 12,    // 减小字体
                            font: "微软雅黑",
                            bold: isHeader
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { 
                        line: 160,    // 减小行距
                        before: 0,
                        after: 0
                    }
                })
            ]
        });
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
            const sizeText = `${size}(${stock})`;
            const separator = index < sizesWithStock.length - 1 ? " " : "";
            
            if (stock < 2) {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: sizeText,
                        color: "FF0000",
                        bold: true,
                        size: 12,    // 减小字体
                        font: "微软雅黑"
                    })
                );
            } else {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: sizeText,
                        color: "000000",
                        size: 12,    // 减小字体
                        font: "微软雅黑"
                    })
                );
            }
            
            if (separator) {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: separator,
                        color: "000000",
                        size: 12,
                        font: "微软雅黑"
                    })
                );
            }
        });

        const cell = new docx.TableCell({
            width: { size: parseInt(width), type: docx.WidthType.PERCENTAGE },
            margins: { 
                top: 30,    // 减小边距
                bottom: 30, 
                left: 30, 
                right: 30 
            },
            shading: { fill: bgColor },
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    children: paragraphChildren,
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { 
                        line: 160,    // 减小行距
                        before: 0,
                        after: 0
                    }
                })
            ]
        });

        return cell;
    }

    // ... 其他方法保持不变 ...

    showResult(data) {
        const separatedProducts = this.separateVandVWProducts(data);
        const vProducts = separatedProducts.vProducts;
        const vwProducts = separatedProducts.vwProducts;
        
        // 计算预估页数
        const totalProducts = data.length;
        const estimatedPages = Math.ceil(totalProducts / 60) + Math.ceil(totalProducts / 80); // 粗略估算
        
        const resultText = '处理完成！共找到 ' + data.length + ' 个产品组合\n' +
                         'V系列: ' + vProducts.length + '个\n' +
                         'VW系列: ' + vwProducts.length + '个\n' +
                         '预估页数: ' + estimatedPages + '页\n\n' +
                         '✅ 优化特性：\n' +
                         '• 最大化纸张利用，减少打印页数\n' +
                         '• 只显示有库存的尺码\n' +
                         '• 库存<2的尺码显示为红色（需要补货）\n' +
                         '• 压缩商品条码和颜色列宽度\n' +
                         '• 增加尺码列显示空间\n' +
                         '• 删除↑符号，同款后续行留空\n' +
                         '• 均衡左右两栏宽度';
        
        document.getElementById('resultText').innerHTML = resultText.replace(/\n/g, '<br>');
        
        document.getElementById('progressArea').style.display = 'none';
        document.getElementById('resultArea').style.display = 'block';
        
        document.getElementById('downloadBtn').onclick = () => this.downloadWordDocument();
    }
}

// ... 其他代码保持不变 ...