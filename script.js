createSizesCell(product, bgColor, width) {
    if (!product) {
        return this.createTableCell("", false, width);
    }

    // 创建带红色格式的尺码文本
    const paragraphChildren = [];
    const sizesWithStock = product.sizes;
    
    sizesWithStock.forEach(([size, stock], index) => {
        const separator = index < sizesWithStock.length - 1 ? " " : "";
        
        if (stock > 2) {
            // 库存>2的尺码用红色
            paragraphChildren.push(
                new docx.TextRun({
                    text: size,
                    color: "FF0000", // 红色
                    size: 14,
                    font: "微软雅黑"
                })
            );
        } else {
            // 正常库存的尺码用黑色
            paragraphChildren.push(
                new docx.TextRun({
                    text: size,
                    color: "000000", // 黑色
                    size: 14,
                    font: "微软雅黑"
                })
            );
        }
        
        // 添加分隔空格
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