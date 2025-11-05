class InventoryReplenishmentGenerator {
    constructor() {
        this.sizeOrder = {
            'XS': 1, 'S': 2, 'M': 3, 'L': 4, 'XL': 5, 'XXL': 6,
            '28': 7, '29': 8, '30': 9, '31': 10, '32': 11, '33': 12,
            '34': 13, '36': 14, '38': 15, 'ONS': 16
        };
        this.doc = null;
        this.processedData = null;
        
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');

        // 文件选择事件
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFileSelect(e.target.files[0]);
            }
        });

        // 拖放事件
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
            
            await this.generateWordDocument(processedData);
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

    // 显示所有列名用于调试
    const availableColumns = Object.keys(data[0]);
    console.log('Excel文件中的列名:', availableColumns);

    // 检查必需列
    const requiredColumns = ['商品条码', 'Color', 'Size', '总库存'];
    const missingColumns = requiredColumns.filter(col => 
        !data[0] || !(col in data[0])
    );
    
    if (missingColumns.length > 0) {
        throw new Error(`缺少必需的列：${missingColumns.join(', ')}`);
    }

    // 过滤有库存的数据，且Color和Size不为空
    const filteredData = data.filter(row => {
        const stock = parseFloat(row['总库存']) || 0;
        const color = row['Color'] ? row['Color'].toString().trim() : '';
        const size = row['Size'] ? row['Size'].toString().trim() : '';
        
        return stock > 0 && color !== '' && size !== '';
    });

    console.log(`过滤后数据: ${filteredData.length} 行（排除库存为0或颜色尺寸为空的行）`);

    if (filteredData.length === 0) {
        throw new Error('没有找到有效的库存数据（请确保有库存>0且颜色和尺寸不为空的数据）');
    }

    // 按商品条码和颜色分组
    const grouped = {};
    filteredData.forEach(row => {
        const productCode = row['商品条码'] ? row['商品条码'].toString().trim() : '';
        const color = row['Color'] ? row['Color'].toString().trim() : '';
        const size = row['Size'] ? row['Size'].toString().trim() : '';
        
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
  
    sortSizes(sizes) {
        const getSizeKey = (size) => {
            const sizeStr = size.toUpperCase();
            if (this.sizeOrder.hasOwnProperty(sizeStr)) {
                return this.sizeOrder[sizeStr];
            }
            if (!isNaN(sizeStr)) {
                return parseInt(sizeStr) + 100; // 数字尺寸排在后面
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

    async generateWordDocument(processedData) {
        const { vProducts, vwProducts } = this.separateVandVWProducts(processedData);
        
        // 这里使用docx.js创建Word文档
        // 由于docx.js在浏览器中使用较复杂，我们使用简化方案
        this.processedData = { vProducts, vwProducts };
    }

    generateOutputContent() {
        const { vProducts, vwProducts } = this.processedData;
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
            
            // 左侧产品
            if (leftProduct) {
                line += this.formatProductLine(leftProduct).padEnd(30);
            } else {
                line += "".padEnd(30);
            }
            
            // 右侧产品
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
        
        // 设置下载按钮事件
        document.getElementById('downloadBtn').onclick = () => this.downloadWordDocument();
    }

    showError(message) {
        document.getElementById('progressArea').style.display = 'none';
        document.getElementById('uploadArea').style.display = 'block';
        alert('错误：' + message);
    }
}

// 全局函数
function resetApp() {
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadArea').style.display = 'block';
    document.getElementById('progressArea').style.display = 'none';
    document.getElementById('resultArea').style.display = 'none';
}

// 初始化应用
let app;
document.addEventListener('DOMContentLoaded', () => {
    app = new InventoryReplenishmentGenerator();
});