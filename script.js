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

    initializeEventListeners() {
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

        // 文件选择事件 - 修复现有的点击事件
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

    async handleFiles(files) {
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

    isExcelFile(file) {
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        const allowedExtensions = ['.xlsx', '.xls'];
        
        return allowedTypes.includes(file.type) || 
               allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    }

    showProgress() {
        const uploadArea = document.getElementById('uploadArea');
        const progressArea = document.getElementById('progressArea');
        
        if (uploadArea) uploadArea.style.display = 'none';
        if (progressArea) progressArea.style.display = 'block';
        
        this.updateProgress('正在读取Excel文件...', 30);
    }

    updateProgress(text, percent) {
        const progressText = document.getElementById('progressText');
        const progressFill = document.getElementById('progressFill');
        
        if (progressText) progressText.textContent = text;
        if (progressFill) progressFill.style.width = percent + '%';
    }

    showResult() {
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

    resetApp() {
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

    async readExcelFile(file) {
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

    processData(data) {
        console.log('🔄 开始处理数据...');
        this.updateProgress('处理库存数据...', 80);
        
        // 这里放置你原有的数据处理逻辑
        // 暂时返回示例数据
        const processed = data.slice(1).map((row, index) => ({
            code: row[0] || `V${index + 1000}`,
            color: row[1] || 'Black',
            size: row[2] || 'M',
            stock: row[3] || 0
        }));
        
        console.log('✅ 数据处理完成，产品数量:', processed.length);
        return processed;
    }

    async downloadWordDocument() {
        try {
            this.updateProgress('生成Word文档...', 90);
            
            // 调用你原有的 generateWordDocument 方法
            if (typeof this.generateWordDocument === 'function') {
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
            } else {
                // 如果没有 generateWordDocument 方法，创建一个简单的示例
                alert('Word文档生成功能待实现');
            }
            
        } catch (error) {
            console.error('❌ 生成Word文档时出错:', error);
            alert('生成Word文档时出错: ' + error.message);
        }
    }

    // 添加你原有的其他方法...
    // generateWordDocument, createProductTablesWithContinuousPagination, 等等
}

// 全局函数，用于HTML中的onclick事件
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