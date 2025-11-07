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

    // ... 事件监听器和其他方法保持不变 ...

    async generateWordDocument() {
        const separatedProducts = this.separateVandVWProducts(this.processedData);
        const vProducts = separatedProducts.vProducts;
        const vwProducts = separatedProducts.vwProducts;
        
        const docChildren = [];
        let currentPage = 1;
        let totalPages = 0;

        // 先计算总页数
        const vTables = this.createProductTablesWithPagination(vProducts, "V");
        const vwTables = this.createProductTablesWithPagination(vwProducts, "VW");
        totalPages = vTables.length + vwTables.length;

        // 处理V系列产品
        if (vProducts.length > 0) {
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "MENS audit list",
                            bold: true,
                            size: 22,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 100 }
                })
            );

            // 添加时间
            const timestamp = new Date().toLocaleString('zh-CN');
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "Time: " + timestamp,
                            size: 12,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { after: 100 }
                })
            );

            vTables.forEach((table, index) => {
                if (index > 0) {
                    docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
                    docChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: "MENS audit list",
                                    bold: true,
                                    size: 22,
                                    font: "微软雅黑"
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 100 }
                        })
                    );
                }
                docChildren.push(table);
                
                // 添加页码
                docChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Page " + currentPage + " of " + totalPages,
                                size: 12,
                                font: "微软雅黑",
                                color: "666666"
                            })
                        ],
                        alignment: docx.AlignmentType.RIGHT,
                        spacing: { after: 100 }
                    })
                );
                currentPage++;
            });
        }

        // 处理VW系列产品
        if (vwProducts.length > 0) {
            docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
            
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "WOMENS audit list",
                            bold: true,
                            size: 22,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 100 }
                })
            );

            // 添加时间
            const timestamp = new Date().toLocaleString('zh-CN');
            docChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "Time: " + timestamp,
                            size: 12,
                            font: "微软雅黑"
                        })
                    ],
                    alignment: docx.AlignmentType.LEFT,
                    spacing: { after: 100 }
                })
            );

            vwTables.forEach((table, index) => {
                if (index > 0) {
                    docChildren.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
                    docChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: "WOMENS audit list",
                                    bold: true,
                                    size: 22,
                                    font: "微软雅黑"
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 100 }
                        })
                    );
                }
                docChildren.push(table);
                
                // 添加页码
                docChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Page " + currentPage + " of " + totalPages,
                                size: 12,
                                font: "微软雅黑",
                                color: "666666"
                            })
                        ],
                        alignment: docx.AlignmentType.RIGHT,
                        spacing: { after: 100 }
                    })
                );
                currentPage++;
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
                            top: 500,  // 减小页边距
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

    // 创建分页表格，每页80行数据
    createProductTablesWithPagination(products, type) {
        const groupedProducts = this.groupProductsByCode(products);
        const tables = [];
        
        // 每页显示80行数据（包含表头）
        const rowsPerPage = 80;
        const dataRowsPerPage = rowsPerPage - 1; // 减去表头行
        
        const totalPages = Math.ceil(groupedProducts.length / dataRowsPerPage);

        console.log(type + "系列：共 " + groupedProducts.length + " 个产品，分 " + totalPages + " 页显示，每页 " + dataRowsPerPage + " 个产品");

        for (let page = 0; page < totalPages; page++) {
            const startIdx = page * dataRowsPerPage;
            const endIdx = startIdx + dataRowsPerPage;
            const pageProducts = groupedProducts.slice(startIdx, endIdx);
            
            const table = this.createProductTable(pageProducts);
            tables.push(table);
        }

        return tables;
    }

    // 创建产品表格 - 紧凑布局显示80行
    createProductTable(products) {
        const productsPerColumn = Math.ceil(products.length / 2);

        const tableRows = [];

        // 表头行 - 紧凑高度
        const headerRow = new docx.TableRow({
            children: [
                this.createTableCell("商品条码", true, "10%"),
                this.createTableCell("颜色", true, "6%"),
                this.createTableCell("尺码", true, "39%"),
                this.createTableCell("商品条码", true, "10%"),
                this.createTableCell("颜色", true, "6%"),
                this.createTableCell("尺码", true, "39%")
            ],
            height: {
                value: 200,  // 减小表头高度
                rule: docx.HeightRule.EXACT
            }
        });
        tableRows.push(headerRow);

        for (let i = 0; i < productsPerColumn; i++) {
            const leftProduct = products[i];
            const rightProduct = products[i + productsPerColumn];
            const bgColor = i % 2 === 0 ? "F8F8F8" : "FFFFFF";

            // 数据行 - 紧凑高度
            const row = new docx.TableRow({
                children: [
                    this.createProductCell(leftProduct, bgColor, "10%"),
                    this.createColorCell(leftProduct, bgColor, "6%"),
                    this.createSizesCell(leftProduct, bgColor, "39%"),
                    this.createProductCell(rightProduct, bgColor, "10%"),
                    this.createColorCell(rightProduct, bgColor, "6%"),
                    this.createSizesCell(rightProduct, bgColor, "39%")
                ],
                height: {
                    value: 180,  // 减小数据行高度
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
            columnWidths: [10, 6, 39, 10, 6, 39],
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
                top: 20,    // 减小边距
                bottom: 20, 
                left: 20, 
                right: 20 
            },
            shading: isHeader ? { fill: "E8E8E8" } : undefined,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: text,
                            size: 12,    // 适当减小字体
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
            const sizeText = size + "(" + stock + ")";
            const separator = index < sizesWithStock.length - 1 ? " " : "";
            
            if (stock < 2) {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: sizeText,
                        color: "FF0000",
                        bold: true,
                        size: 12,    // 适当减小字体
                        font: "微软雅黑"
                    })
                );
            } else {
                paragraphChildren.push(
                    new docx.TextRun({
                        text: sizeText,
                        color: "000000",
                        size: 12,    // 适当减小字体
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
                top: 20,    // 减小边距
                bottom: 20, 
                left: 20, 
                right: 20 
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
        
        const totalProducts = data.length;
        const estimatedPages = Math.ceil(totalProducts / 79); // 79行数据
        
        const resultText = '处理完成！共找到 ' + data.length + ' 个产品组合\n' +
                         'V系列: ' + vProducts.length + '个\n' +
                         'VW系列: ' + vwProducts.length + '个\n' +
                         '预估页数: ' + estimatedPages + '页\n\n' +
                         '优化特性：\n' +
                         '• 每页80行，最大化利用纸张\n' +
                         '• 优化列宽：商品条码10%，颜色6%，尺码39%\n' +
                         '• 紧凑布局设计\n' +
                         '• 完整页码统计\n' +
                         '• 只显示有库存的尺码\n' +
                         '• 库存<2的尺码显示为红色';
        
        document.getElementById('resultText').innerHTML = resultText.replace(/\n/g, '<br>');
        
        document.getElementById('progressArea').style.display = 'none';
        document.getElementById('resultArea').style.display = 'block';
        
        document.getElementById('downloadBtn').onclick = () => this.downloadWordDocument();
    }
}

// ... 其他代码保持不变 ...