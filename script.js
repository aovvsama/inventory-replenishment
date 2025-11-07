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

    // 创建连续分页的表格
    createProductTablesWithContinuousPagination(products, type) {
        const groupedProducts = this.groupProductsByCode(products);
        const tables = [];
        
        // 每页显示的产品数量（根据实际调整，每页20行，每行左右各1个产品）
        const productsPerPage = 40; // 每页左右两栏各20个产品
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

    // ... 其他方法保持不变 ...
}