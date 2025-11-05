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
            
            reader.onerror = () => reject(new Error('读取文件失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    processData(data) {
        if (!data || data.length === 0) {
            throw new Error('Excel文件中没有数据');
        }

        console.log('原始数据第一行:', data[0]);
        console.log('所有列名:', Object.keys(data[0]));

        // 智能列名映射
        const columnMap = this.detectColumns(data[0]);
        console.log('检测到的列映射:', columnMap);

        // 检查必需列
        const missingColumns = [];
        if (!columnMap.productCode) missingColumns.push('商品条码');
        if (!columnMap.color) missingColumns.push('Color');
        if (!columnMap.size) missingColumns.push('Size');
        if (!columnMap.stock) missingColumns.push('总库存');

        if (missingColumns.length > 0) {
            throw new Error(`缺少必需的列：${missingColumns.join(', ')}\n\n检测到的列名：${Object.keys(data[0]).join(', ')}`);
        }

        // 过滤有库存的数据，且Color和Size不为空
        const filteredData = data.filter(row => {
            const stock = parseFloat(row[columnMap.stock]) || 0;
            const color = row[columnMap.color] ? row[columnMap.color].toString().trim() : '';
            const size = row[columnMap.size] ? row[columnMap.size].toString().trim() : '';
            
            return stock > 0 && color !== '' && size !== '';
        });

        console.log(`过滤后数据: ${filteredData.length} 行`);

        if (filteredData.length === 0) {
            throw new Error('没有找到有效的库存数据（请确保有库存>0且颜色和尺寸不为空的数据）');
        }

        // 按商品条码和颜色分组
        const grouped = {};
        filteredData.forEach(row => {
            const productCode = row[columnMap.productCode] ? row[columnMap.productCode].toString().trim() : '';
            const color = row[columnMap.color] ? row[columnMap.color].toString().trim() : '';
            const size = row[columnMap.size] ? row[columnMap.size].toString().trim() : '';
            
            if (productCode && color) {
                const key = `${productCode}_${color}`;
                if (!grouped[key]) {
                    grouped[key] = {
                        productCode: productCode,
                        color: color,
                        sizes: new Set()
                    };
                }
                if (size) {
                    grouped[key].sizes.add(size);
                }
            }
        });

        // 转换并排序尺寸
        const processedData = Object.values(grouped).map(item => ({
            productCode: item.productCode,
            color: item.color,
            sizes: this.sortSizes(Array.from(item.sizes))
        }));

        console.log(`处理后产品数量: ${processedData.length}`);
        console.log('处理后的数据:', processedData);

        return processedData;
    }

    detectColumns(firstRow) {
        const columns = Object.keys(firstRow);
        const mapping = {
            productCode: null,
            color: null,
            size: null,
            stock: null
        };

        // 定义可能的列名匹配
        const patterns = {
            productCode: ['商品条码', '条码', 'sku', 'code', '商品编码'],
            color: ['color', '颜色', 'colour', '色号'],
            size: ['size', '尺码', '尺寸', '规格'],
            stock: ['总库存', '库存', 'stock', 'quantity', '库存数']
        };

        // 为每个类型找到匹配的列
        Object.keys(patterns).forEach(key => {
            for (const pattern of patterns[key]) {
                const found = columns.find(col => 
                    col.toLowerCase().includes(pattern.toLowerCase())
                );
                if (found) {
                    mapping[key] = found;
                    break;
                }
            }
        });

        return mapping;
    }

    sortSizes(sizes) {
        const getSizeKey = (size) => {
            const sizeStr = size.toUpperCase();
            if (this.sizeOrder.hasOwnProperty(sizeStr)) {
                return this.sizeOrder[sizeStr];
            }
            if (!isNaN(sizeStr)) {
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
            if (item.productCode.toString().startsWith('VW')) {
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

    generateOutputContent() {
        const { vProducts, vwProducts } = this.separateVandVWProducts(this.processedData);
        let content = "库存补货清单\n\n";
        const timestamp = new Date().toLocaleString('zh-CN');

        if (vProducts.length > 0) {
            content += "MENS audit list\n";
            content += "=".repeat(50) + "\n";
            const groupedV = this.groupProductsByCode(vProducts);
            content += this.formatProductsForText(groupedV);
            content += "\n\n";
        }

        if (vwProducts.length > 0) {
            content += "WOMENS audit list\n";
            content += "=".repeat(50) + "\n";
            const groupedVW = this.groupProductsByCode(vwProducts);
            content += this.formatProductsForText(groupedVW);
        }

        content += `\n生成时间: ${timestamp}\n`;
        return content;
    }

    formatProductsForText(products) {
        let content = "";
        const chunkSize = Math.ceil(products.length / 2);
        
        for (let i = 0; i < chunkSize; i++) {
            const leftProduct = products[i];
            const rightProduct = products[i + chunkSize];
            
            let line = "";
            
            if (leftProduct) {
                line += this.formatProductLine(leftProduct).padEnd(30);
            } else {
                line += "".padEnd(30);
            }
            
            if (rightProduct) {
                line += this.formatProductLine(rightProduct);
            }
            
            content += line + "\n";
        }
        
        return content;
    }

    formatProductLine(product) {
        const code = product.showCode ? product.productCode : "";
        const color = product.color || "";
        const sizes = this.formatSizesText(product.sizes, 20);
        return `${code}\t${color}\t${sizes}`;
    }

    downloadWordDocument() {
        const content = this.generateOutputContent();
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        saveAs(blob, `补货清单_${timestamp}.txt`);
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
        document.getElementById('resultText').textContent = 
            `处理完成！共找到 ${data.length} 个产品（V系列: ${vProducts.length}个, VW系列: ${vwProducts.length}个）`;
        
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