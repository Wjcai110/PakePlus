document.addEventListener('DOMContentLoaded', function() {
    // 定义全局变量
    const canvas = document.getElementById('template-canvas');
    const importTemplateBtn = document.getElementById('import-template');
    const importTemplateConfigBtn = document.getElementById('import-template-config');
    const addFieldBtn = document.getElementById('add-field');
    const deleteFieldBtn = document.getElementById('delete-field');
    const saveTemplateBtn = document.getElementById('save-template');
    const importBtn = document.getElementById('import-data');
    const exportBtn = document.getElementById('export-data');
    const printBtn = document.getElementById('print');
    const printBackgroundCheckbox = document.getElementById('print-background');
    const dataModal = document.getElementById('data-modal');
    const templateSizeModal = document.getElementById('template-size-modal');
    const closeBtns = document.querySelectorAll('.close');
    const confirmDataBtn = document.getElementById('confirm-data');
    const confirmSizeBtn = document.getElementById('confirm-size');
    // const dataInput = document.getElementById('data-input'); // 已移除JSON输入功能
    const fileInput = document.getElementById('file-input');
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');
    const templateWidthInput = document.getElementById('template-width');
    const templateHeightInput = document.getElementById('template-height');
    const applyPropertiesBtn = document.getElementById('apply-properties');
    const zoomInBtn = document.getElementById('zoom-in');
    const zoomOutBtn = document.getElementById('zoom-out');
    const propertiesToggle = document.getElementById('properties-toggle');
    const propertiesPanel = document.querySelector('.properties-panel');
    const mainContent = document.querySelector('.main-content');
    const templateSizeSelect = document.getElementById('template-size-select');
    
    // IndexedDB 相关变量
    let db;
    const DB_NAME = 'templateEditorDB';
    const DB_VERSION = 1;
    const STORE_NAME = 'templateData';
    const TEMPLATE_KEY = 'templateEditorData';
    
    // 模板和字段相关变量
    let templateImage = null;
    let fields = [];
    let selectedField = null;
    let isDragging = false;
    let isResizing = false;
    let startX, startY, startLeft, startTop;
    let templateWidth = 297; // 默认A4横版宽度，单位毫米
    let templateHeight = 210; // 默认A4横版高度，单位毫米
    let currentZoom = 1; // 当前缩放比例
    let templateData = []; // 用于存储导入的数据
    let csvData = null; // 用于存储导入的CSV数据
    
    // 获取设备像素比
    const devicePixelRatio = window.devicePixelRatio || 1;
    
    // 对齐辅助线
    let alignmentGuides = {
        vertical: null,
        horizontal: null
    };
    
    // 对齐吸附阈值（毫米）
    const snapThreshold = 2;
    
    // 属性面板状态
    let isPropertiesPanelVisible = true;
    
    // 动态计算实际显示比例，而不依赖固定DPI
    function getActualDisplayRatio() {
        // 方法1：使用画布实际尺寸比例
        const mmToPixelRatio = canvas.offsetWidth / templateWidth;
        return mmToPixelRatio;
    }
    
    // 结合设备像素比的计算
    function getEffectiveDPI() {
        const baseDPI = 96; // CSS标准DPI
        const deviceRatio = window.devicePixelRatio || 1;
        const zoomFactor = canvas.offsetWidth / templateWidth / (templateWidth / 210); // 假设A4宽度210mm
        return baseDPI * deviceRatio * zoomFactor;
    }
    
    // 直接使用比例，避免DPI概念
    function mmToPxAccurate(mm) {
        const displayRatio = canvas.offsetWidth / templateWidth;
        return Math.round(mm * displayRatio);
    }
    
    // 单位转换函数 - 毫米转屏幕像素（用于显示）
    function mmToPx(mm) {
        const mmToPixelRatio = canvas.offsetWidth / templateWidth;
        return Math.round(mm * mmToPixelRatio);
    }
    
    // 单位转换函数 - 屏幕像素转毫米（用于保存坐标）
    function pxToMm(px) {
        const mmToPixelRatio = canvas.offsetWidth / templateWidth;
        return Math.round((px / mmToPixelRatio) * 100) / 100;
    }
    
    // 初始化属性面板状态
    function initializePropertiesPanel() {
        // 从本地存储加载属性面板状态
        const savedState = localStorage.getItem('propertiesPanelState');
        if (savedState === 'hidden') {
            isPropertiesPanelVisible = false;
            propertiesPanel.classList.add('hidden');
            mainContent.classList.add('expanded');
            document.querySelector('.container').classList.add('expanded');
            const toggleText = propertiesToggle.querySelector('.toggle-text');
            toggleText.textContent = '显示属性';
        }
    }
    
    // 切换属性面板显示状态
    function togglePropertiesPanel() {
        isPropertiesPanelVisible = !isPropertiesPanelVisible;
        propertiesPanel.classList.toggle('hidden');
        mainContent.classList.toggle('expanded');
        document.querySelector('.container').classList.toggle('expanded');
        
        // 更新切换按钮文本
        const toggleText = propertiesToggle.querySelector('.toggle-text');
        toggleText.textContent = isPropertiesPanelVisible ? '隐藏属性' : '显示属性';
        
        // 保存状态到本地存储
        localStorage.setItem('propertiesPanelState', isPropertiesPanelVisible ? 'visible' : 'hidden');
    }
    
    // 绑定属性面板切换按钮事件
    propertiesToggle.addEventListener('click', togglePropertiesPanel);
    

    
    // 初始化页面
    initIndexedDB()
        .then(() => {
            initializePage();
            initializePropertiesPanel();
            enhancePropertiesPanel();
        })
        .catch(error => {
            console.error('初始化IndexedDB失败:', error);
            // 如果IndexedDB初始化失败，仍然继续初始化页面
            initializePage();
            initializePropertiesPanel();
            enhancePropertiesPanel();
        });
    
    // 初始化函数
    function initializePage() {
        // 设置初始模板大小为A4纸张大小（以毫米为单位）
        canvas.style.width = templateWidth + 'mm';
        canvas.style.height = templateHeight + 'mm';
        
        // 从本地存储加载数据
        if (db) {
            loadFromLocalStorage();
        } else {
            // 如果IndexedDB未初始化，直接使用localStorage
            loadFromLocalStorageFallback();
        }
        
        // 绑定事件处理函数
        bindEventListeners();
    }
    
    // 绑定事件处理函数
    function bindEventListeners() {
        // 导入模板按钮
        importTemplateBtn.addEventListener('click', importTemplate);
        
        // 导入模板配置按钮
        importTemplateConfigBtn.addEventListener('click', importTemplateConfig);
        
        // 上传按钮
        const uploadBtn = document.querySelector('.upload-btn');
        if (uploadBtn) {
            uploadBtn.addEventListener('click', function() {
                fileInput.click(); // 触发隐藏的文件输入框
            });
        }
        
        // 添加字段按钮
        addFieldBtn.addEventListener('click', addField);
        
        // 删除字段按钮
        deleteFieldBtn.addEventListener('click', deleteField);
        
        // 保存模板按钮
        saveTemplateBtn.addEventListener('click', saveTemplate);
        
        // 放大按钮
        zoomInBtn.addEventListener('click', zoomIn);
        
        // 缩小按钮
        zoomOutBtn.addEventListener('click', zoomOut);
        
        // 导入数据按钮
        importBtn.addEventListener('click', function() {
            resetDataModal();
            dataModal.style.display = 'block';
        });
        
        // 导出数据按钮
        exportBtn.addEventListener('click', exportData);
        
        // 打印按钮
        printBtn.addEventListener('click', printTemplate);
        

        
        // 确认尺寸按钮
        confirmSizeBtn.addEventListener('click', confirmTemplateSize);
        
        // 关闭所有模态框按钮
        closeBtns.forEach(btn => {
            btn.addEventListener('click', function() {
            dataModal.style.display = 'none';
                templateSizeModal.style.display = 'none';
            });
        });
        

        
        // 确认数据按钮
        confirmDataBtn.addEventListener('click', confirmData);
        

        
        // 属性应用按钮
        applyPropertiesBtn.addEventListener('click', applyProperties);
        
        // 点击模态框外部关闭模态框
        window.addEventListener('click', (e) => {
            if (e.target === dataModal) {
                dataModal.style.display = 'none';
            }
            if (e.target === templateSizeModal) {
                templateSizeModal.style.display = 'none';
            }
        });
        
        // 字段类型变更监听
        document.getElementById('field-type').addEventListener('change', function() {
            const imageFieldContainer = document.getElementById('image-field-container');
            const fieldDefaultValueContainer = document.getElementById('field-default-value').parentNode;
            
            if (this.value === 'image') {
                fieldDefaultValueContainer.style.display = 'none';
                imageFieldContainer.style.display = 'block';
            } else {
                fieldDefaultValueContainer.style.display = 'block';
                imageFieldContainer.style.display = 'none';
            }
        });
        
        // 图片浏览按钮点击事件
        document.getElementById('browse-image-btn').addEventListener('click', function() {
            document.getElementById('field-image-select').click();
        });
        
        // 图片选择变更事件
        document.getElementById('field-image-select').addEventListener('change', function(e) {
            if (this.files && this.files[0]) {
                const file = this.files[0];
                const reader = new FileReader();
                
                // 保存文件名
                const fileName = file.name;
                
                reader.onload = function(e) {
                    const imageDataUrl = e.target.result; // base64数据
                    
                    // 更新选中字段的图片数据
                    if (selectedField) {
                        const fieldIndex = fields.findIndex(f => f.id === selectedField.id);
                        if (fieldIndex !== -1) {
                            // 保存原始文件名（仅用于显示）
                            fields[fieldIndex].imagePath = fileName;
                            // 保存base64数据用于实际显示
                            fields[fieldIndex].imageDataUrl = imageDataUrl;
                            // 立即应用更改
                            applyProperties();
                        }
                    }
                };
                
                reader.onerror = function() {
                    alert('图片文件读取失败，请重新选择！');
                    console.error('图片文件读取失败');
                };
                
                reader.readAsDataURL(file);
            }
        });
        
        // 复制Base64按钮点击事件
        document.getElementById('copy-base64-btn').addEventListener('click', function() {
            if (selectedField) {
                const fieldIndex = fields.findIndex(f => f.id === selectedField.id);
                if (fieldIndex !== -1 && fields[fieldIndex].imageDataUrl) {
                    // 复制Base64数据到剪贴板
                    navigator.clipboard.writeText(fields[fieldIndex].imageDataUrl)
                        .then(() => {
                            alert('Base64数据已复制到剪贴板！');
                        })
                        .catch(err => {
                            console.error('复制失败:', err);
                            alert('复制失败，请重试！');
                        });
                } else {
                    alert('没有可用的图片数据！请先选择图片。');
                }
            }
        });
        
        // 键盘方向键微调
        document.addEventListener('keydown', function(e) {
            // 确保有选中的字段，并且没有在输入框中
            if (selectedField && !['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)) {
                // 获取字段对象
                const field = fields.find(f => f.id === selectedField.id);
                if (!field) return;
                
                const stepSize = 0.5; // 每次移动0.5毫米
                
                let xChange = 0;
                let yChange = 0;
                
                // 确定移动方向
                switch (e.key) {
                    case 'ArrowLeft':
                        xChange = -stepSize;
                        break;
                    case 'ArrowRight':
                        xChange = stepSize;
                        break;
                    case 'ArrowUp':
                        yChange = -stepSize;
                        break;
                    case 'ArrowDown':
                        yChange = stepSize;
                        break;
                    default:
                        return; // 其他键不处理
                }
                
                // 更新字段位置数据（毫米）
                field.x += xChange;
                field.y += yChange;
                
                // 更新DOM元素位置（像素）
                selectedField.style.left = mmToPx(field.x) + 'px';
                selectedField.style.top = mmToPx(field.y) + 'px';
                
                // 更新属性面板
                document.getElementById('field-x').value = parseFloat(field.x.toFixed(1));
                document.getElementById('field-y').value = parseFloat(field.y.toFixed(1));
                
                // 阻止页面滚动
                e.preventDefault();
            }
        });
        
        // 文件上传
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            handleFileUpload(e);
            // 显示所选文件名
            document.querySelector('.file-name').textContent = file.name || '未选择任何文件';
        }
    });
        
        // 标签切换
        tabButtons.forEach(button => {
            button.addEventListener('click', () => {
                // 移除所有标签的活动状态
                tabButtons.forEach(btn => btn.classList.remove('active'));
                // 隐藏所有内容面板
                tabContents.forEach(content => content.style.display = 'none');
                
                // 激活当前标签和内容
                button.classList.add('active');
                const tabId = button.getAttribute('data-tab');
                document.getElementById(tabId).style.display = 'block';
            });
        });
        
        // 为画布添加点击事件，取消字段选择
        canvas.addEventListener('click', function(e) {
            if (e.target === canvas) {
                if (selectedField) {
                    selectedField.classList.remove('selected');
                    selectedField = null;
                }
            }
        });
        
        // 为画布添加字段鼠标按下事件
        canvas.addEventListener('mousedown', function(e) {
            // 如果点击的是字段元素或其子元素
            if (e.target.classList.contains('field') || 
                e.target.closest('.field')) {
                let fieldElement = e.target.classList.contains('field') ? 
                                  e.target : e.target.closest('.field');
                handleFieldMouseDown(e);
            }
        });
    }
    
    // 导入模板函数
    function importTemplate() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        
        input.onchange = function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(event) {
                    // 创建一个新的图像对象
                    const img = new Image();
                    img.src = event.target.result;
                    
                    img.onload = function() {
                        // 清空画布
                        clearCanvas();
                        
                        // 计算图片的实际尺寸（毫米）
                        const widthInMm = pxToMm(img.width);
                        const heightInMm = pxToMm(img.height);
                        
                        // 不直接设置模板尺寸为图片尺寸，而是显示尺寸设置模态框
                        // 让用户决定是否使用图片尺寸或保留当前尺寸
                        
                        // 创建背景图像元素
                        templateImage = document.createElement('img');
                        templateImage.src = img.src;
                        templateImage.className = 'template-image';
                        templateImage.style.width = '100%';
                        templateImage.style.height = '100%';
                        templateImage.style.position = 'absolute';
                        templateImage.style.top = '0';
                        templateImage.style.left = '0';
                        templateImage.style.zIndex = '0';
                        
                        // 将图像添加到画布
                        canvas.appendChild(templateImage);
                        
                        // 移除占位消息
                        const placeholder = canvas.querySelector('.placeholder-message');
                        if (placeholder) {
                            placeholder.style.display = 'none';
                        }
                        
                        // 显示尺寸设置模态框，并预填充图片尺寸
                        templateWidthInput.value = widthInMm;
                        templateHeightInput.value = heightInMm;
                        templateSizeModal.style.display = 'block';
                    };
                    
                    img.onerror = function() {
                        alert('图片加载失败，请检查文件格式是否正确！');
                        console.error('模板图片加载失败');
                    };
                };
                reader.readAsDataURL(file);
            }
        };
        
        input.click();
        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 导入模板配置函数
    function importTemplateConfig() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json';
        
        input.onchange = function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(event) {
                    try {
                        // 解析JSON数据
                        const templateConfig = JSON.parse(event.target.result);
                        
                        // 检查JSON格式是否正确
                        if (templateConfig.fields && Array.isArray(templateConfig.fields)) {
                            // 默认保留当前背景图片
                            const savedTemplateImage = templateImage;
                            
                            // 清除所有字段，但保留背景图片
                            clearFieldsOnly();
                            
                            // 应用模板尺寸（如果有）
                            if (templateConfig.width && templateConfig.height) {
                                templateWidth = templateConfig.width;
                                templateHeight = templateConfig.height;
                                canvas.style.width = templateWidth + 'mm';
                                canvas.style.height = templateHeight + 'mm';
                            }
                            
                            // 加载打印偏移值（如果有）
                            if (templateConfig.printOffset) {
                                // 创建全局变量保存打印偏移值
                                window.savedPrintOffset = {
                                    x: templateConfig.printOffset.x || 0,
                                    y: templateConfig.printOffset.y || 0
                                };
                                console.log('已加载打印偏移值:', window.savedPrintOffset);
                            }
                            
                            // 如果配置中包含背景图片且当前没有背景图片，或用户选择使用配置中的背景图片
                            if (templateConfig.backgroundImage && (!savedTemplateImage || confirm('是否使用配置文件中的背景图片替换当前图片？'))) {
                                const img = new Image();
                                img.src = templateConfig.backgroundImage;
                                
                                img.onload = function() {
                                    // 创建背景图像元素
                                    templateImage = document.createElement('img');
                                    templateImage.src = img.src;
                                    templateImage.className = 'template-image';
                                    templateImage.style.width = '100%';
                                    templateImage.style.height = 'auto';
                                    templateImage.style.position = 'absolute';
                                    templateImage.style.top = '0';
                                    templateImage.style.left = '0';
                                    templateImage.style.zIndex = '0';
                                    
                                    // 将图像添加到画布
                                    canvas.appendChild(templateImage);
                                    
                                    // 创建字段元素
                                    createFieldsAfterImageLoad();
                                };
                                
                                img.onerror = function() {
                                    console.warn('配置文件中的背景图片加载失败');
                                    alert('配置文件中的背景图片加载失败，将跳过背景图片加载。');
                                    // 直接创建字段元素
                                    createFieldsAfterImageLoad();
                                };
                            } else if (savedTemplateImage) {
                                // 恢复之前的背景图片
                                templateImage = savedTemplateImage;
                                
                                // 如果背景图片不在DOM中，重新添加
                                if (!canvas.contains(templateImage)) {
                                canvas.appendChild(templateImage);
                            }
                            
                                // 直接创建字段元素
                                createFieldsAfterImageLoad();
                            } else {
                                // 没有背景图片，直接创建字段元素
                                createFieldsAfterImageLoad();
                            }
                            
                            // 移除占位消息
                            const placeholder = canvas.querySelector('.placeholder-message');
                            if (placeholder) {
                                placeholder.style.display = 'none';
                            }
                            
                            // 提示导入成功
                            alert('模板配置导入成功！');
                            
                            // 内部函数：创建字段元素
                            function createFieldsAfterImageLoad() {
                                // 加载字段
                                fields = templateConfig.fields;
                                
                                // 创建字段元素
                                fields.forEach(field => {
                                    createFieldElement(field);
                                });
                            }
                        } else {
                            alert('模板配置格式错误！');
                        }
                    } catch (e) {
                        alert('解析JSON失败！请确保文件格式正确。');
                        console.error(e);
                    }
                };
                reader.readAsText(file);
            }
        };
        
        input.click();
        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 仅清除字段，保留背景图片的函数
    function clearFieldsOnly() {
        // 移除所有字段元素
        document.querySelectorAll('.field').forEach(field => {
            field.remove();
        });
        
        // 清空字段数组
        fields = [];
        selectedField = null;
        
        // 注意：此函数不会删除背景图片，templateImage变量在此保持不变
    }
    
    // 添加字段函数
    function addField() {
        // 创建唯一ID
        const fieldId = 'field-' + Date.now();
        
        // 创建字段对象（使用毫米单位）
    const field = {
        id: fieldId,
        name: '字段' + (fields.length + 1),
        type: 'text',
        x: 20, // 毫米 - 你想要的X坐标
        y: 20, // 毫米 - 你想要的Y坐标
        width: 25, // 毫米 - 你想要的宽度
        height: 3, // 毫米 - 你想要的高度
        fontSize: 12,
        fontFamily: 'SimSun',
        fontWeight: 'normal', // 添加字体粗细属性
        defaultValue: '',
        imagePath: '' // 新增：图片路径
    };
        
        // 添加到字段数组
        fields.push(field);
        
        // 创建字段DOM元素
        createFieldElement(field);
        
        // 选中新添加的字段
        selectField(document.getElementById(fieldId));
        

        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 创建字段DOM元素
    function createFieldElement(field) {
        // 创建字段元素
        const fieldElement = document.createElement('div');
        fieldElement.id = field.id;
        fieldElement.className = 'field';
        
        // 如果是图片类型，添加额外的样式类
        if (field.type === 'image') {
            fieldElement.classList.add('image-field');
        }
        
        // 考虑缩放因素设置位置
        fieldElement.style.left = mmToPx(field.x) + 'px';
        fieldElement.style.top = mmToPx(field.y) + 'px';
        fieldElement.style.width = mmToPx(field.width) + 'px';
        fieldElement.style.height = mmToPx(field.height) + 'px'; // 应用高度属性
        fieldElement.style.fontSize = field.fontSize + 'pt';
        fieldElement.style.fontFamily = field.fontFamily;
        fieldElement.style.fontWeight = field.fontWeight || 'normal';
        fieldElement.style.fontStyle = field.fontStyle || 'normal';
        fieldElement.style.textDecoration = field.textDecoration || 'none';
        
        // 创建字段标签
        const fieldLabel = document.createElement('div');
        fieldLabel.className = 'field-label';
        fieldLabel.textContent = field.name;
        
        // 根据字段类型处理内容
        if (field.type === 'image' && (field.imageDataUrl || field.imagePath)) {
            // 对于图片类型，添加一个img元素
            const img = document.createElement('img');
            
            // 优先使用base64数据，避免文件路径问题
            if (field.imageDataUrl) {
                img.src = field.imageDataUrl;
            } else if (field.imagePath) {
                // 如果只有文件路径，尝试加载，但添加错误处理
                img.src = field.imagePath;
                img.onerror = function() {
                    console.warn('图片加载失败，清除无效路径:', field.imagePath);
                    // 清除无效的图片路径
                    field.imagePath = '';
                    field.imageDataUrl = '';
                    // 显示占位符
                    this.style.display = 'none';
                    const placeholder = document.createElement('div');
                    placeholder.textContent = '图片加载失败';
                    placeholder.style.cssText = 'color: #999; text-align: center; padding: 10px; border: 1px dashed #ccc;';
                    this.parentNode.appendChild(placeholder);
                };
            }
            
            img.alt = field.name;
            img.style.maxWidth = '100%';
            img.style.maxHeight = '100%';
            fieldElement.appendChild(img);
            
            // 添加图片路径提示
            if (field.imagePath && !field.imageDataUrl) {
                const pathHint = document.createElement('div');
                pathHint.className = 'image-path-hint';
                pathHint.textContent = field.imagePath;
                fieldElement.appendChild(pathHint);
            }
        } else {
            // 对于其他类型，使用文本内容
            const contentSpan = document.createElement('span');
            contentSpan.className = 'field-content';
            contentSpan.textContent = field.defaultValue || field.name;
            // 应用样式
            contentSpan.style.fontWeight = field.fontWeight || 'normal';
            contentSpan.style.fontStyle = field.fontStyle || 'normal';
            contentSpan.style.textDecoration = field.textDecoration || 'none';
            fieldElement.appendChild(contentSpan);
        }
        
        fieldElement.appendChild(fieldLabel);
        
        // 添加调整大小的控制点
        const positions = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
        positions.forEach(pos => {
            const handle = document.createElement('div');
            handle.className = `resize-handle ${pos}`;
            handle.setAttribute('data-position', pos);
            
            // 添加控制点的鼠标按下事件
            handle.addEventListener('mousedown', function(e) {
                e.stopPropagation(); // 阻止事件冒泡，避免触发字段的拖动
                handleResizeStart(e, fieldElement, pos);
            });
            
            fieldElement.appendChild(handle);
        });
        
        console.log("创建字段元素，添加了控制点:", positions.length);
        
        // 添加点击事件以选择字段
        fieldElement.addEventListener('click', function(e) {
            selectField(this);
            e.stopPropagation();
        });
        
        // 添加双击事件以编辑字段内容
        fieldElement.addEventListener('dblclick', function(e) {
            e.stopPropagation();
            // 移除只有非图片字段才能双击编辑的限制
            handleFieldDoubleClick(this);
        });
        
        // 将字段元素添加到画布
        canvas.appendChild(fieldElement);
    }
    
    // 选择字段函数
    function selectField(fieldElement) {
        // 取消之前选中的字段
        if (selectedField) {
            selectedField.classList.remove('selected');
        }
        
        // 选中当前字段
        selectedField = fieldElement;
        if (selectedField) {
            selectedField.classList.add('selected');
            console.log("字段已选中:", selectedField.id);
            
            // 获取字段对象
            const field = fields.find(f => f.id === selectedField.id);
            
            if (field) {
            // 更新属性面板
            document.getElementById('field-name').value = field.name;
            document.getElementById('field-type').value = field.type;
                document.getElementById('field-x').value = field.x; // 毫米
                document.getElementById('field-y').value = field.y; // 毫米
                document.getElementById('field-width').value = field.width; // 毫米
                document.getElementById('field-height').value = field.height; // 毫米
                document.getElementById('field-font-size').value = field.fontSize;
                
                // 处理字体选择和自定义字体
                const fontFamilySelect = document.getElementById('field-font-family');
                const customFontContainer = document.getElementById('custom-font-container');
                const customFontInput = document.getElementById('field-custom-font');
                
                // 检查是否是预定义字体之一
                const isStandardFont = Array.from(fontFamilySelect.options).some(option => 
                    option.value !== 'custom' && option.value === field.fontFamily
                );
                
                if (isStandardFont) {
                    fontFamilySelect.value = field.fontFamily;
                    customFontContainer.style.display = 'none';
                } else {
                    fontFamilySelect.value = 'custom';
                    customFontContainer.style.display = 'block';
                    customFontInput.value = field.fontFamily;
                }
                
                // 设置加粗状态
                const boldBtn = document.getElementById('field-font-bold');
                document.getElementById('field-font-bold').checked = (field.fontWeight === 'bold');
                document.getElementById('field-font-italic').checked = (field.fontStyle === 'italic');
                document.getElementById('field-font-underline').checked = (field.textDecoration === 'underline');
                
                // 设置斜体状态
                const italicBtn = document.getElementById('field-font-italic');
                if (field.fontStyle === 'italic') {
                    italicBtn.classList.add('active');
                } else {
                    italicBtn.classList.remove('active');
                }
                
                // 设置下划线状态
                const underlineBtn = document.getElementById('field-font-underline');
                if (field.textDecoration === 'underline') {
                    underlineBtn.classList.add('active');
                } else {
                    underlineBtn.classList.remove('active');
                }
                

                
                document.getElementById('field-default-value').value = field.defaultValue || '';
                
                // 处理图片字段相关UI显示
                const imageFieldContainer = document.getElementById('image-field-container');
                const fieldDefaultValueContainer = document.getElementById('field-default-value').parentNode;
                
                // 根据字段类型显示/隐藏相应控件
                if (field.type === 'image') {
                    fieldDefaultValueContainer.style.display = 'none';
                    imageFieldContainer.style.display = 'block';
                    
                    // 处理图片字段显示
                    if (field.imagePath && !field.imageDataUrl) {
                        // 如果只有路径但没有数据URL
                        const pathDisplay = document.createElement('div');
                        pathDisplay.textContent = "文件名: " + field.imagePath;
                        pathDisplay.style.marginTop = '5px';
                        pathDisplay.style.fontSize = '12px';
                        
                        // 清除之前可能添加的路径显示
                        const existingPathDisplay = imageFieldContainer.querySelector('.path-display');
                        if (existingPathDisplay) {
                            existingPathDisplay.remove();
                        }
                        
                        pathDisplay.className = 'path-display';
                        imageFieldContainer.appendChild(pathDisplay);
                    }
                } else {
                    fieldDefaultValueContainer.style.display = 'block';
                    imageFieldContainer.style.display = 'none';
                }
                
                selectedField.style.fontSize = field.fontSize + 'pt';
                selectedField.style.fontFamily = field.fontFamily;
                selectedField.style.fontWeight = field.fontWeight;
                
                // 确保控制点可见
                const resizeHandles = selectedField.querySelectorAll('.resize-handle');
                resizeHandles.forEach(handle => {
                    handle.style.display = 'block';
                });
                
                // 同步内容样式
                const contentSpan = selectedField.querySelector('.field-content');
                if (contentSpan) {
                    contentSpan.style.fontWeight = field.fontWeight || 'normal';
                    contentSpan.style.fontStyle = field.fontStyle || 'normal';
                    contentSpan.style.textDecoration = field.textDecoration || 'none';
                }
            }
        }
    }
    
    // 删除字段函数
    function deleteField() {
        if (selectedField) {
            // 从数组中移除
            fields = fields.filter(f => f.id !== selectedField.id);
            
            // 从DOM中移除
            selectedField.remove();
            
            // 重置选中状态
            selectedField = null;
        }
        

        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 应用字体样式函数
    function applyFontStyles() {
        if (!selectedField) return;
        
        // 获取字段对象
        const field = fields.find(f => f.id === selectedField.id);
        if (!field) return;
        
        // 获取字体样式状态
        const isBold = document.getElementById('field-font-bold').checked;
        const isItalic = document.getElementById('field-font-italic').checked;
        const hasUnderline = document.getElementById('field-font-underline').checked;
        
        // 更新字段对象
        field.fontWeight = isBold ? 'bold' : 'normal';
        field.fontStyle = isItalic ? 'italic' : 'normal';
        field.textDecoration = hasUnderline ? 'underline' : 'none';
        
        // 更新DOM元素样式
        selectedField.style.fontWeight = field.fontWeight;
        selectedField.style.fontStyle = field.fontStyle;
        selectedField.style.textDecoration = field.textDecoration;
        
        // 只作用于内容span
        const contentSpan = selectedField.querySelector('.field-content');
        if (contentSpan) {
            contentSpan.style.fontWeight = field.fontWeight;
            contentSpan.style.fontStyle = field.fontStyle;
            contentSpan.style.textDecoration = field.textDecoration;
        }
        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 应用属性函数
    function applyProperties() {
        // 检查是否有选中的字段
        if (selectedField) {
            // 获取字段在数组中的索引
            const index = fields.findIndex(f => f.id === selectedField.id);
            
            if (index !== -1) {
                // 获取属性值
                const name = document.getElementById('field-name').value;
                const type = document.getElementById('field-type').value;
                const x = parseFloat(document.getElementById('field-x').value); // 毫米
                const y = parseFloat(document.getElementById('field-y').value); // 毫米
                const width = parseFloat(document.getElementById('field-width').value); // 毫米
                const height = parseFloat(document.getElementById('field-height').value); // 毫米
                const fontSize = parseInt(document.getElementById('field-font-size').value);
                
                // 处理字体选择
                let fontFamily = document.getElementById('field-font-family').value;
                if (fontFamily === 'custom') {
                    fontFamily = document.getElementById('field-custom-font').value.trim();
                    if (!fontFamily) {
                        fontFamily = 'SimSun'; // 默认字体
                    }
                }
                
                // 处理字体加粗
                const isBold = document.getElementById('field-font-bold').checked;
                const fontWeight = isBold ? 'bold' : 'normal';
                
                // 处理斜体
                const isItalic = document.getElementById('field-font-italic').checked;
                const fontStyle = isItalic ? 'italic' : 'normal';
                
                // 处理下划线
                const hasUnderline = document.getElementById('field-font-underline').checked;
                const textDecoration = hasUnderline ? 'underline' : 'none';
                

                
                const defaultValue = document.getElementById('field-default-value').value;
                
                // 获取当前图片路径（如果有）
                let imagePath = fields[index].imagePath || '';
                let imageDataUrl = fields[index].imageDataUrl || '';
                
                // 更新字段对象（使用毫米值）
                fields[index].name = name;
                fields[index].type = type;
                fields[index].x = x;
                fields[index].y = y;
                fields[index].width = width;
                fields[index].height = height;
                fields[index].fontSize = fontSize;
                fields[index].fontFamily = fontFamily;
                fields[index].fontWeight = fontWeight;
                fields[index].fontStyle = fontStyle;
                fields[index].textDecoration = textDecoration;

                fields[index].defaultValue = defaultValue;
                fields[index].imagePath = imagePath;
                fields[index].imageDataUrl = imageDataUrl;
                
                // 更新DOM元素（转换为像素）
                selectedField.style.left = mmToPx(x) + 'px';
                selectedField.style.top = mmToPx(y) + 'px';
                selectedField.style.width = mmToPx(width) + 'px';
                selectedField.style.height = mmToPx(height) + 'px';
                selectedField.style.fontSize = fontSize + 'pt';
                selectedField.style.fontFamily = fontFamily;
                selectedField.style.fontWeight = fontWeight;
                selectedField.style.fontStyle = fontStyle;
                selectedField.style.textDecoration = textDecoration;

                
                // 清除旧内容
                selectedField.innerHTML = '';
                
                // 根据字段类型更新内容
                if (type === 'image') {
                    selectedField.classList.add('image-field');
                    
                    if (imageDataUrl || imagePath) {
                        // 创建图片元素
                        const img = document.createElement('img');
                        
                        // 优先使用base64数据，避免文件路径问题
                        if (imageDataUrl) {
                            img.src = imageDataUrl;
                        } else if (imagePath) {
                            // 如果只有文件路径，尝试加载，但添加错误处理
                            img.src = imagePath;
                            img.onerror = function() {
                                console.warn('图片加载失败，清除无效路径:', imagePath);
                                // 清除无效的图片路径
                                fields[index].imagePath = '';
                                fields[index].imageDataUrl = '';
                                // 显示占位符
                                this.style.display = 'none';
                                const placeholder = document.createElement('div');
                                placeholder.textContent = '图片加载失败';
                                placeholder.style.cssText = 'color: #999; text-align: center; padding: 10px; border: 1px dashed #ccc;';
                                this.parentNode.appendChild(placeholder);
                            };
                        }
                        
                        img.alt = name;
                        img.style.maxWidth = '100%';
                        img.style.maxHeight = '100%';
                        selectedField.appendChild(img);
                        
                        // 如果只有文件路径但没有数据URL，显示文件名
                        if (imagePath && !imageDataUrl) {
                            const pathHint = document.createElement('div');
                            pathHint.className = 'image-path-hint';
                            pathHint.textContent = imagePath;
                            selectedField.appendChild(pathHint);
                        }
                    }
                } else {
                    selectedField.classList.remove('image-field');
                    // 更新文本内容
                    const contentSpan = document.createElement('span');
                    contentSpan.className = 'field-content';
                    contentSpan.textContent = defaultValue || name;
                    contentSpan.style.fontWeight = fontWeight;
                    contentSpan.style.fontStyle = fontStyle;
                    contentSpan.style.textDecoration = textDecoration;
                    selectedField.appendChild(contentSpan);
                }
                
                // 重新添加标签元素
                const fieldLabel = document.createElement('div');
                fieldLabel.className = 'field-label';
                fieldLabel.textContent = name;
                selectedField.appendChild(fieldLabel);
                
                // 重新添加调整大小的控制点
                const positions = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
                positions.forEach(pos => {
                    const handle = document.createElement('div');
                    handle.className = `resize-handle ${pos}`;
                    handle.setAttribute('data-position', pos);
                    
                    handle.addEventListener('mousedown', function(e) {
                        e.stopPropagation();
                        handleResizeStart(e, selectedField, pos);
                    });
                    
                    selectedField.appendChild(handle);
                });
            }
        }
        
        saveToLocalStorage();
    }
    
    // 保存模板函数
    function saveTemplate() {
        // 确保每个字段对象都包含必要属性
        const fieldsToSave = fields.map(field => {
            // 创建字段的副本，避免修改原始对象
            const fieldCopy = { ...field };
            
            // 确保有默认属性
            fieldCopy.defaultValue = field.defaultValue || '';
            fieldCopy.imagePath = field.imagePath || '';
            
            // 处理图片字段 - 不保存base64数据到配置文件
            if (field.type === 'image') {
                // 保存图片路径但不保存base64数据
                // 确保移除大型的base64数据，避免配置文件过大
                delete fieldCopy.imageDataUrl;
            }
            
            return fieldCopy;
        });
        
        // 获取当前的打印偏移值
        const printXOffset = parseFloat(document.querySelector('#xOffsetInput')?.value || '0');
        const printYOffset = parseFloat(document.querySelector('#yOffsetInput')?.value || '0');
        
        const templateData = {
            fields: fieldsToSave,
            width: templateWidth,
            height: templateHeight,
            printOffset: {
                x: printXOffset,
                y: printYOffset
            },
            // 添加背景图片数据
            backgroundImage: templateImage ? templateImage.src : null,
            // 添加版本信息
            version: '1.0.0'
        };
        
        // 将数据编码为JSON并转换为URI编码格式
        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(templateData));
        
        // 创建下载链接
        const downloadAnchor = document.createElement('a');
        downloadAnchor.setAttribute("href", dataStr);
        
        // 设置文件名
        const now = new Date();
        const timestamp = now.getFullYear() + 
                      ("0" + (now.getMonth() + 1)).slice(-2) + 
                      ("0" + now.getDate()).slice(-2) + 
                      ("0" + now.getHours()).slice(-2) + 
                      ("0" + now.getMinutes()).slice(-2);
                      
        downloadAnchor.setAttribute("download", "模板配置_" + timestamp + ".json");
        
        // 模拟点击下载
        document.body.appendChild(downloadAnchor);
        downloadAnchor.click();
        
        // 清理
        document.body.removeChild(downloadAnchor);
        
        console.log("模板已保存，尺寸: " + templateWidth + "mm x " + templateHeight + "mm" + 
                  (templateData.printOffset.x !== 0 || templateData.printOffset.y !== 0 ?
                  '，偏移值: X=' + templateData.printOffset.x + 'mm, Y=' + templateData.printOffset.y + 'mm' : ''));
    }
    
    // 重置数据导入模态框状态
    function resetDataModal() {
        // 重置文件输入
        fileInput.value = '';
        // 重置文件名显示
        document.querySelector('.file-name').textContent = '未选择任何文件';
        // 重置CSV数据
        csvData = null;
    }
    
    // 处理文件上传
    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        // 检查文件类型
        const isCSV = file.type === 'text/csv' || file.name.endsWith('.csv');
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
        
        if (!isCSV && !isExcel) {
            alert('请上传Excel格式的文件！');
            fileInput.value = '';
            return;
        }
        
        if (isCSV) {
            // 处理CSV文件
            const reader = new FileReader();
            reader.onload = function(event) {
                try {
                    const content = event.target.result;
                    parseCSV(content);
                } catch (e) {
                    console.error('解析CSV文件时出错:', e);
                    alert('解析CSV文件失败，请确保文件格式正确！');
                }
            };
            
            reader.onerror = function(event) {
                console.error('读取文件时出错:', event.target.error);
                alert('读取文件失败，请检查文件是否有效！');
            };
            
            reader.readAsText(file, 'GBK');
        } else {
            // 处理Excel文件
            handleExcelFile(file);
        }
    }
    
    // 处理Excel文件
    function handleExcelFile(file) {
        // 显示文件名
        document.querySelector('.file-name').textContent = file.name || '未选择任何文件';
        
        // 处理Excel文件
        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 获取第一个工作表
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // 将工作表转换为JSON
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                
                if (jsonData.length < 2) {
                    alert('Excel文件格式不正确，至少需要包含表头行和一行数据！');
                    return;
                }
                
                // 获取表头行
                const headers = jsonData[0].map(header => String(header).trim());
                
                // 获取数据行
                const dataRows = jsonData.slice(1).map(row => 
                    row.map(cell => cell === undefined ? '' : String(cell).trim())
                );
                
                // 添加调试输出
                console.log('解析到的表头:', headers);
                console.log('解析到的数据行:', dataRows);
                console.log('Excel样本数据:');
                for (let i = 0; i < Math.min(3, dataRows.length); i++) {
                    const rowObj = {};
                    for (let j = 0; j < Math.min(headers.length, dataRows[i].length); j++) {
                        rowObj[headers[j]] = dataRows[i][j];
                    }
                    console.log(`第${i+1}行:`, rowObj);
                }
                
                // 存储数据
                csvData = {
                    headers: headers,
                    rows: dataRows
                };
                
                // 生成预览表格
                generateFilePreview(headers, dataRows);
                
                console.log('Excel文件解析成功，共', dataRows.length, '行数据');
            } catch (e) {
                console.error('解析Excel文件时出错:', e);
                alert('解析Excel文件失败，请确保文件格式正确！');
            }
        };
        
        reader.onerror = function(event) {
            console.error('读取文件时出错:', event.target.error);
            alert('读取文件失败，请检查文件是否有效！');
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // 生成文件预览表格
    function generateFilePreview(headers, dataRows) {
        const previewContainer = document.getElementById('file-preview');
        if (!previewContainer) return;
        
        // 限制预览行数（最多显示前5行）
        const maxPreviewRows = Math.min(5, dataRows.length);
        
        let previewHTML = `
            <div class="preview-info">
                <p><strong>文件预览</strong> - 共 ${dataRows.length} 行数据，显示前 ${maxPreviewRows} 行</p>
            </div>
            <div class="preview-table-container">
                <table class="preview-table">
                    <thead>
                        <tr>
        `;
        
        // 添加表头
        headers.forEach(header => {
            previewHTML += `<th>${escapeHTML(header)}</th>`;
        });
        
        previewHTML += `
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        // 添加数据行 - 增强图片字段处理
        for (let i = 0; i < maxPreviewRows; i++) {
            previewHTML += '<tr>';
            for (let j = 0; j < headers.length; j++) {
                const cellValue = dataRows[i] && dataRows[i][j] ? dataRows[i][j] : '';
                
                // 检查是否为图片字段（BASE64数据）
                if (cellValue && (cellValue.startsWith('data:image/') || 
                    (cellValue.length > 100 && /^[A-Za-z0-9+/=]+$/.test(cellValue)))) {
                    // 显示图片预览
                    const imgSrc = cellValue.startsWith('data:image/') ? cellValue : `data:image/jpeg;base64,${cellValue}`;
                    previewHTML += `<td class="image-cell">
                        <img src="${imgSrc}" alt="图片预览" class="preview-image" 
                             style="max-width: 80px; max-height: 60px; object-fit: contain;" 
                             onerror="this.style.display='none'; this.nextSibling.style.display='block';">
                        <span style="display: none; color: #999; font-size: 11px;">图片加载失败</span>
                    </td>`;
                } else {
                    // 普通文本数据
                    previewHTML += `<td>${escapeHTML(cellValue)}</td>`;
                }
            }
            previewHTML += '</tr>';
        }
        
        previewHTML += `
                    </tbody>
                </table>
            </div>
        `;
        
        if (dataRows.length > maxPreviewRows) {
            previewHTML += `<p class="preview-note">还有 ${dataRows.length - maxPreviewRows} 行数据未显示...</p>`;
        }
        
        previewContainer.innerHTML = previewHTML;
    }
    
    // 解析CSV数据
    function parseCSV(content) {
        // 处理CSV内容，支持特殊字符和换行
        const processCSV = (str) => {
            const arr = [];
            let quote = false;  // 是否在引号内
            let row = [];
            let cell = '';
            
            for (let i = 0; i < str.length; i++) {
                const currChar = str[i];
                const nextChar = str[i + 1];
                
                // 处理引号内的内容
                if (currChar === '"') {
                    if (quote && nextChar === '"') {
                        cell += '"';
                        i++;
                    } else {
                        quote = !quote;
                    }
                    continue;
                }
                
                // 处理逗号分隔符
                if (currChar === ',' && !quote) {
                    row.push(cell.trim());
                    cell = '';
                    continue;
                }
                
                // 处理换行符
                if (currChar === '\r' && nextChar === '\n' && !quote) {
                    row.push(cell.trim());
                    arr.push(row);
                    row = [];
                    cell = '';
                    i++;
                    continue;
                }
                
                // 处理单独的换行符
                if ((currChar === '\r' || currChar === '\n') && !quote) {
                    row.push(cell.trim());
                    arr.push(row);
                    row = [];
                    cell = '';
                    continue;
                }
                
                // 正常字符
                cell += currChar;
            }
            
            // 处理最后一个单元格和行
            if (cell !== '') {
                row.push(cell.trim());
            }
            if (row.length > 0) {
                arr.push(row);
            }
            
            return arr;
        };
        
        // 使用BOM检测UTF-8编码，并移除BOM标记
        let csvString = content;
        if (csvString.charCodeAt(0) === 0xFEFF) {
            csvString = csvString.substr(1); // 移除BOM
        }
        
        try {
            // 解析CSV字符串
            const csvArray = processCSV(csvString);
            
            if (csvArray.length < 2) {
                alert('CSV文件格式不正确，至少需要包含表头行和一行数据！');
                return;
            }
            
            // 获取表头行
            const headers = csvArray[0].map(header => header.trim());
            
            // 获取数据行
            const dataRows = csvArray.slice(1);
            
            // 添加调试输出
            console.log('解析到的表头:', headers);
            console.log('解析到的数据行:', dataRows);
            console.log('CSV样本数据:');
            for (let i = 0; i < Math.min(3, dataRows.length); i++) {
                const rowObj = {};
                for (let j = 0; j < Math.min(headers.length, dataRows[i].length); j++) {
                    rowObj[headers[j]] = dataRows[i][j];
                }
                console.log(`第${i+1}行:`, rowObj);
            }
            
            // 存储CSV数据
                csvData = {
                    headers: headers,
                    rows: dataRows
                };
                
                // 生成预览表格
                generateFilePreview(headers, dataRows);
                
                console.log('CSV文件解析成功，共', dataRows.length, '行数据');
        } catch (error) {
            console.error('CSV解析错误:', error);
            alert('CSV文件解析失败: ' + error.message);
        }
    }
    

    

    
    // 辅助函数：转义HTML特殊字符，防止XSS攻击
    function escapeHTML(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }
    
    // 确认数据函数
    function confirmData() {
        try {
            // 检查是否已有字段
            if (fields.length === 0) {
                alert('请先添加字段！');
                return;
            }
            
            // 初始化空模板数据
            templateData = [];
            let missingFields = [];
            
            // 由于移除了标签页，直接处理文件上传
            // 检查是否已上传CSV文件
            if (!csvData || !csvData.headers || !csvData.rows || csvData.rows.length === 0) {
                alert('请先上传有效的Excel文件！');
                return;
            }
            
            // 获取数据头和所有行
            const headers = csvData.headers;
            const allRows = csvData.rows;
            
            console.log('导入数据处理 - 表头:', headers);
            console.log('导入数据处理 - 行数:', allRows.length);
            
            // 遍历每一行，创建模板数据对象
            for (let rowIndex = 0; rowIndex < allRows.length; rowIndex++) {
                const currentRow = allRows[rowIndex];
                const rowData = {};
                
                // 确保headers和currentRow长度匹配
                const maxLength = Math.min(headers.length, currentRow.length);
                
                for (let i = 0; i < maxLength; i++) {
                    if (headers[i]) {
                        // 去除可能的空格和引号
                        const headerName = headers[i].trim().replace(/^["']|["']$/g, '');
                        const cellValue = currentRow[i] ? currentRow[i].trim().replace(/^["']|["']$/g, '') : '';
                        
                        rowData[headerName] = cellValue;
                        
                        if (rowIndex < 3) {
                            console.log(`数据行 ${rowIndex}, 字段 ${headerName}:`, cellValue);
                        }
                    }
                }
                
                // 将当前行数据添加到模板数据数组
                templateData.push(rowData);
            }
            
            console.log('导入数据完成，共 ' + templateData.length + ' 行');
            console.log('首行数据样本:', templateData[0]);
            
            // 使用第一组数据检查缺失字段
            const firstRowData = templateData[0];
            // 检查是否缺少必要字段
            fields.forEach(field => {
                if (!field.name) return; // 跳过没有名称的字段
                
                let fieldFound = false;
                
                // 尝试多种匹配方式
                if (field.name in firstRowData) fieldFound = true;
                else if (field.name.trim() in firstRowData) fieldFound = true;
                else if (field.name.toLowerCase() in firstRowData) fieldFound = true;
                else if (field.name.toUpperCase() in firstRowData) fieldFound = true;
                
                console.log('字段检查 ' + field.name + ':', {
                    fieldName: field.name,
                    inDataRow: fieldFound,
                    hasDefaultValue: !!field.defaultValue,
                    defaultValue: field.defaultValue
                });
                
                // 如果字段没有默认值，检查是否在导入数据中存在
                if (!field.defaultValue && !fieldFound) {
                    missingFields.push(field.name);
                }
            });
            
            // 显示导入结果
            if (missingFields.length > 0) {
                alert('警告：以下字段在数据中缺失：' + missingFields.join(', ') + 
                      '\n请检查数据格式或为这些字段设置默认值。');
            } else {
                alert('成功导入 ' + templateData.length + ' 组数据！');
                dataModal.style.display = 'none';
            }
            
            // 保存到本地存储
            saveToLocalStorage();
            
        } catch (error) {
            console.error('确认数据时出错:', error);
            alert('处理数据时出错: ' + error.message);
        }
    }
    

    

    
    // 打印模板函数
    function printTemplate() {
        if (!templateData || templateData.length === 0) {
            alert('请先导入数据！');
            return;
        }
        
        // 创建打印选项对话框
        const printOptionsDialog = document.createElement('div');
        printOptionsDialog.className = 'modal';
        printOptionsDialog.style.display = 'block';
        printOptionsDialog.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
        
        // 创建对话框内容
        const dialogContent = document.createElement('div');
        dialogContent.className = 'modal-content';
        dialogContent.style.width = '400px';
        dialogContent.style.padding = '25px';
        dialogContent.style.backgroundColor = 'white';
        dialogContent.style.borderRadius = '8px';
        dialogContent.style.boxShadow = '0 4px 20px rgba(0,0,0,0.15)';
        dialogContent.style.position = 'fixed';
        dialogContent.style.top = '20%'; // 修改这里，从50%改为45%
        dialogContent.style.left = '50%';
        dialogContent.style.transform = 'translate(-50%, -50%)';
        dialogContent.style.fontFamily = 'Microsoft YaHei, sans-serif';
        
        // 添加标题
        const title = document.createElement('h3');
        title.textContent = '打印选项';
        title.style.margin = '0 0 20px 0';
        title.style.color = '#333';
        title.style.fontSize = '18px';
        title.style.fontWeight = '500';
        title.style.borderBottom = '1px solid #eee';
        title.style.paddingBottom = '10px';
        
        // 添加打印范围选择
        const rangeContainer = document.createElement('div');
        rangeContainer.style.marginBottom = '20px';
        
        const rangeLabel = document.createElement('label');
        rangeLabel.textContent = '打印范围';
        rangeLabel.style.display = 'block';
        rangeLabel.style.marginBottom = '10px';
        rangeLabel.style.color = '#555';
        rangeLabel.style.fontSize = '14px';
        rangeLabel.style.fontWeight = '500';
        
        const radioContainer = document.createElement('div');
        radioContainer.style.display = 'flex';
        radioContainer.style.alignItems = 'center';
        radioContainer.style.gap = '15px';
        
        const allOption = document.createElement('input');
        allOption.type = 'radio';
        allOption.name = 'printRange';
        allOption.value = 'all';
        allOption.checked = true;
        allOption.id = 'printAll';
        allOption.style.margin = '0';
        
        const allLabel = document.createElement('label');
        allLabel.textContent = '所有数据 (' + templateData.length + '组)';
        allLabel.htmlFor = 'printAll';
        allLabel.style.margin = '0';
        allLabel.style.color = '#666';
        allLabel.style.fontSize = '14px';
        
        const currentOption = document.createElement('input');
        currentOption.type = 'radio';
        currentOption.name = 'printRange';
        currentOption.value = 'current';
        currentOption.id = 'printCurrent';
        currentOption.style.margin = '0';
        
        const currentLabel = document.createElement('label');
        currentLabel.textContent = '指定范围';
        currentLabel.htmlFor = 'printCurrent';
        currentLabel.style.margin = '0';
        currentLabel.style.color = '#666';
        currentLabel.style.fontSize = '14px';
        
        radioContainer.appendChild(allOption);
        radioContainer.appendChild(allLabel);
        radioContainer.appendChild(currentOption);
        radioContainer.appendChild(currentLabel);
        
        rangeContainer.appendChild(rangeLabel);
        rangeContainer.appendChild(radioContainer);
        
        // 添加范围输入
        const rangeInputContainer = document.createElement('div');
        rangeInputContainer.style.marginBottom = '20px';
        rangeInputContainer.style.display = 'none';
        rangeInputContainer.style.padding = '15px';
        rangeInputContainer.style.backgroundColor = '#f8f9fa';
        rangeInputContainer.style.borderRadius = '4px';
        
        const rangeInputRow = document.createElement('div');
        rangeInputRow.style.display = 'flex';
        rangeInputRow.style.alignItems = 'center';
        rangeInputRow.style.gap = '10px';
        
        const fromLabel = document.createElement('label');
        fromLabel.textContent = '从';
        fromLabel.style.color = '#666';
        fromLabel.style.fontSize = '14px';
        
        const fromInput = document.createElement('input');
        fromInput.type = 'number';
        fromInput.min = '1';
        fromInput.max = templateData.length.toString();
        fromInput.value = '1';
        fromInput.style.width = '60px';
        fromInput.style.padding = '6px 8px';
        fromInput.style.border = '1px solid #ddd';
        fromInput.style.borderRadius = '4px';
        fromInput.style.fontSize = '14px';
        
        const toLabel = document.createElement('label');
        toLabel.textContent = '至';
        toLabel.style.color = '#666';
        toLabel.style.fontSize = '14px';
        
        const toInput = document.createElement('input');
        toInput.type = 'number';
        toInput.min = '1';
        toInput.max = templateData.length.toString();
        toInput.value = templateData.length.toString();
        toInput.style.width = '60px';
        toInput.style.padding = '6px 8px';
        toInput.style.border = '1px solid #ddd';
        toInput.style.borderRadius = '4px';
        toInput.style.fontSize = '14px';
        
        rangeInputRow.appendChild(fromLabel);
        rangeInputRow.appendChild(fromInput);
        rangeInputRow.appendChild(toLabel);
        rangeInputRow.appendChild(toInput);
        
        rangeInputContainer.appendChild(rangeInputRow);
        
        // 显示背景选项
        const bgContainer = document.createElement('div');
        bgContainer.style.marginBottom = '20px';
        bgContainer.style.display = 'flex';
        bgContainer.style.alignItems = 'center';
        bgContainer.style.gap = '8px';
        
        const bgCheckbox = document.createElement('input');
        bgCheckbox.type = 'checkbox';
        bgCheckbox.id = 'printBg';
        bgCheckbox.checked = false; // 修改默认值为false
        bgCheckbox.style.margin = '0';
        
        const bgLabel = document.createElement('label');
        bgLabel.textContent = '打印背景图片';
        bgLabel.htmlFor = 'printBg';
        bgLabel.style.margin = '0';
        bgLabel.style.color = '#666';
        bgLabel.style.fontSize = '14px';
        
        bgContainer.appendChild(bgCheckbox);
        bgContainer.appendChild(bgLabel);
        
        // 添加打印偏移设置
        const offsetContainer = document.createElement('div');
        offsetContainer.style.marginBottom = '20px';
        
        const offsetTitle = document.createElement('div');
        offsetTitle.textContent = '打印偏移调整';
        offsetTitle.style.marginBottom = '10px';
        offsetTitle.style.color = '#555';
        offsetTitle.style.fontSize = '14px';
        offsetTitle.style.fontWeight = '500';
        
        const offsetInputsContainer = document.createElement('div');
        offsetInputsContainer.style.display = 'grid';
        offsetInputsContainer.style.gridTemplateColumns = '1fr 1fr';
        offsetInputsContainer.style.gap = '15px';
        
        const xOffsetContainer = document.createElement('div');
        const xOffsetLabel = document.createElement('label');
        xOffsetLabel.textContent = '左右偏移 (mm)';
        xOffsetLabel.style.display = 'block';
        xOffsetLabel.style.marginBottom = '5px';
        xOffsetLabel.style.color = '#666';
        xOffsetLabel.style.fontSize = '14px';
        
        const xOffsetInput = document.createElement('input');
        xOffsetInput.type = 'number';
        xOffsetInput.step = '0.5';
        xOffsetInput.id = 'xOffsetInput';
        xOffsetInput.value = window.savedPrintOffset ? window.savedPrintOffset.x.toString() : '0';
        xOffsetInput.style.width = '100%';
        xOffsetInput.style.padding = '8px 12px';
        xOffsetInput.style.border = '1px solid #ddd';
        xOffsetInput.style.borderRadius = '4px';
        xOffsetInput.style.fontSize = '14px';
        xOffsetInput.style.boxSizing = 'border-box';
        
        xOffsetContainer.appendChild(xOffsetLabel);
        xOffsetContainer.appendChild(xOffsetInput);
        
        const yOffsetContainer = document.createElement('div');
        const yOffsetLabel = document.createElement('label');
        yOffsetLabel.textContent = '上下偏移 (mm)';
        yOffsetLabel.style.display = 'block';
        yOffsetLabel.style.marginBottom = '5px';
        yOffsetLabel.style.color = '#666';
        yOffsetLabel.style.fontSize = '14px';
        
        const yOffsetInput = document.createElement('input');
        yOffsetInput.type = 'number';
        yOffsetInput.step = '0.5';
        yOffsetInput.id = 'yOffsetInput';
        yOffsetInput.value = window.savedPrintOffset ? window.savedPrintOffset.y.toString() : '0';
        yOffsetInput.style.width = '100%';
        yOffsetInput.style.padding = '8px 12px';
        yOffsetInput.style.border = '1px solid #ddd';
        yOffsetInput.style.borderRadius = '4px';
        yOffsetInput.style.fontSize = '14px';
        yOffsetInput.style.boxSizing = 'border-box';
        
        yOffsetContainer.appendChild(yOffsetLabel);
        yOffsetContainer.appendChild(yOffsetInput);
        
        offsetInputsContainer.appendChild(xOffsetContainer);
        offsetInputsContainer.appendChild(yOffsetContainer);
        
        offsetContainer.appendChild(offsetTitle);
        offsetContainer.appendChild(offsetInputsContainer);
        
        // 添加按钮容器
        const buttonContainer = document.createElement('div');
        buttonContainer.style.display = 'flex';
        buttonContainer.style.justifyContent = 'flex-end';
        buttonContainer.style.gap = '12px';
        buttonContainer.style.marginTop = '25px';
        
        // 取消按钮
        const cancelButton = document.createElement('button');
        cancelButton.textContent = '取消';
        cancelButton.style.padding = '8px 20px';
        cancelButton.style.backgroundColor = '#f5f5f5';
        cancelButton.style.color = '#666';
        cancelButton.style.border = '1px solid #ddd';
        cancelButton.style.borderRadius = '4px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.style.fontSize = '14px';
        cancelButton.style.transition = 'all 0.3s';
        
        cancelButton.addEventListener('mouseover', () => {
            cancelButton.style.backgroundColor = '#e8e8e8';
        });
        
        cancelButton.addEventListener('mouseout', () => {
            cancelButton.style.backgroundColor = '#f5f5f5';
        });
        
        // 确认按钮
        const confirmButton = document.createElement('button');
        confirmButton.textContent = '打印';
        confirmButton.style.padding = '8px 20px';
        confirmButton.style.backgroundColor = '#4a90e2';
        confirmButton.style.color = 'white';
        confirmButton.style.border = 'none';
        confirmButton.style.borderRadius = '4px';
        confirmButton.style.cursor = 'pointer';
        confirmButton.style.fontSize = '14px';
        confirmButton.style.transition = 'background-color 0.3s';
        
        confirmButton.addEventListener('mouseover', () => {
            confirmButton.style.backgroundColor = '#357abd';
        });
        
        confirmButton.addEventListener('mouseout', () => {
            confirmButton.style.backgroundColor = '#4a90e2';
        });
        
        buttonContainer.appendChild(cancelButton);
        buttonContainer.appendChild(confirmButton);
        
        // 添加所有元素到对话框
        dialogContent.appendChild(title);
        dialogContent.appendChild(rangeContainer);
        dialogContent.appendChild(rangeInputContainer);
        dialogContent.appendChild(bgContainer);
        dialogContent.appendChild(offsetContainer);
        dialogContent.appendChild(buttonContainer);
        
        printOptionsDialog.appendChild(dialogContent);
        document.body.appendChild(printOptionsDialog);
        
        // 事件处理
        currentOption.addEventListener('change', () => {
            if (currentOption.checked) {
                rangeInputContainer.style.display = 'block';
            }
        });
        
        allOption.addEventListener('change', () => {
            if (allOption.checked) {
                rangeInputContainer.style.display = 'none';
            }
        });
        
        cancelButton.addEventListener('click', () => {
            document.body.removeChild(printOptionsDialog);
        });
        
        confirmButton.addEventListener('click', () => {
            // 获取打印选项
            const printAll = allOption.checked;
            const startIndex = parseInt(fromInput.value) - 1;
            const endIndex = parseInt(toInput.value) - 1;
            const shouldShowBackground = bgCheckbox.checked;
            const xOffset = parseFloat(xOffsetInput.value) || 0;
            const yOffset = parseFloat(yOffsetInput.value) || 0;
            
            // 更新全局保存的偏移值
            window.savedPrintOffset = { x: xOffset, y: yOffset };
            console.log('已更新打印偏移值:', window.savedPrintOffset);
            
            // 验证范围
            if (!printAll && (startIndex > endIndex || startIndex < 0 || endIndex >= templateData.length)) {
                alert('请输入有效的打印范围！');
                return;
            }
            
            // 确定要打印的数据范围
            let dataToPrint;
            if (printAll) {
                dataToPrint = templateData;
            } else {
                dataToPrint = templateData.slice(startIndex, endIndex + 1);
            }
            
            // 移除对话框
            document.body.removeChild(printOptionsDialog);
            
            // 直接执行打印，不显示预览窗口
            const printHTML = createPrintHTML(dataToPrint, shouldShowBackground, xOffset, yOffset);
            const printFrame = document.createElement('iframe');
            printFrame.style.position = 'fixed';
            printFrame.style.right = '0';
            printFrame.style.bottom = '0';
            printFrame.style.width = '0';
            printFrame.style.height = '0';
            printFrame.style.border = '0';
            document.body.appendChild(printFrame);
            
            const printDoc = printFrame.contentDocument || printFrame.contentWindow.document;
            printDoc.write(printHTML);
            printDoc.close();
            
            // 等待内容加载完成后直接打印
            printFrame.onload = function() {
                setTimeout(function() {
                    printFrame.contentWindow.focus();
                    printFrame.contentWindow.print();
                    // 打印完成后移除iframe
                    setTimeout(function() {
                        document.body.removeChild(printFrame);
                    }, 1000);
                }, 500);
            };
        });
    }
    
    // 执行打印函数
    function performPrint(dataToPrint, shouldShowBackground, xOffset, yOffset) {
        console.log('执行打印 - 背景显示状态:', shouldShowBackground);
        
        // 创建HTML头部内容
        const htmlHead = `<!DOCTYPE html>
<html>
<head>
    <title>套打打印</title>
    <meta charset="UTF-8">
    <style>
        @page {
            size: ${templateWidth}mm ${templateHeight}mm;
            margin: 0;
        }
        body {
            margin: 0;
            padding: 0;
            overflow: hidden;
        }
        .print-page {
            position: relative;
            width: ${templateWidth}mm;
            height: ${templateHeight}mm;
            margin: 0;
            padding: 0;
            page-break-after: always;
            overflow: visible;
        }
        .print-page:last-child {
            page-break-after: avoid;
        }
        .field-element {
            position: absolute;
            overflow: visible;
            white-space: normal;
            z-index: 10;
        }
        .background-image {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 0;
        }
        @media print {
            body {
                margin: 0;
                padding: 0;
            }
            .print-page {
                display: block !important;
                visibility: visible !important;
                overflow: visible !important;
            }
            .field-element {
                display: block !important;
                visibility: visible !important;
                overflow: visible !important;
            }
        }
    </style>
</head>
<body>`;

        // 创建HTML页面内容
        let htmlBody = '';
        
        // 为每组数据创建一个打印页面
        dataToPrint.forEach((data, index) => {
            htmlBody += '<div class="print-page">';
            
            // 添加背景图片（如果选择显示）
            if (shouldShowBackground && templateImage) {
                console.log('添加背景图片到打印页面');
                htmlBody += `<img class="background-image" src="${templateImage.src}" alt="模板背景" />`;
            } else {
                console.log('不显示背景图片 - shouldShowBackground:', shouldShowBackground, 'templateImage存在:', !!templateImage);
            }
            
            // 添加所有字段
            fields.forEach(field => {
                // 尝试多种可能的字段名匹配方式
                let fieldValue = '';
                
                // 1. 直接匹配字段名
                if (data[field.name] !== undefined) {
                    fieldValue = data[field.name];
                }
                // 2. 尝试去除引号和空格后匹配
                else if (data[field.name.trim()] !== undefined) {
                    fieldValue = data[field.name.trim()];
                }
                // 3. 尝试转为小写后匹配
                else if (data[field.name.toLowerCase()] !== undefined) {
                    fieldValue = data[field.name.toLowerCase()];
                }
                // 4. 尝试转为大写后匹配
                else if (data[field.name.toUpperCase()] !== undefined) {
                    fieldValue = data[field.name.toUpperCase()];
                }
                // 5. 如果都没匹配到，使用默认值
                else {
                    fieldValue = field.defaultValue || '';
                }
                
                // 处理图片字段
                if (field.type === 'image' && fieldValue) {
                    // 如果是BASE64数据，创建img元素
                    if (fieldValue.startsWith('data:image/') || fieldValue.length > 100) {
                        htmlBody += `<div class="field-element" 
                                style="left: ${field.x + xOffset}mm; 
                                       top: ${field.y + yOffset}mm; 
                                       width: ${field.width}mm; 
                                       height: ${field.height}mm;">
                               <img src="${fieldValue}" alt="${field.name}" style="width: 100%; height: 100%; object-fit: contain;" />
                           </div>`;
                    } else {
                        // 如果不是BASE64，可能是文件路径，显示为普通文本
                        htmlBody += `<div class="field-element" 
                                style="left: ${field.x + xOffset}mm; 
                                       top: ${field.y + yOffset}mm; 
                                       width: ${field.width}mm; 
                                       height: ${field.height}mm;
                                       font-size: ${field.fontSize}pt; 
                                       font-family: ${field.fontFamily};
                                       font-weight: ${field.fontWeight || 'normal'};
                                       font-style: ${field.fontStyle || 'normal'};
                                       text-decoration: ${field.textDecoration || 'none'};">
                               ${fieldValue}
                           </div>`;
                    }
                } else {
                    // 普通字段处理
                    htmlBody += `<div class="field-element" 
                            style="left: ${field.x + xOffset}mm; 
                                   top: ${field.y + yOffset}mm; 
                                   width: ${field.width}mm; 
                                   height: ${field.height}mm;
                                   font-size: ${field.fontSize}pt; 
                                   font-family: ${field.fontFamily};
                                   font-weight: ${field.fontWeight || 'normal'};
                                   font-style: ${field.fontStyle || 'normal'};
                                   text-decoration: ${field.textDecoration || 'none'};">
                           ${fieldValue}
                       </div>`;
                }
            });
            
            htmlBody += '</div>';
        });
        
        // 创建HTML尾部和脚本
        const htmlFoot = `<script>
        document.addEventListener('DOMContentLoaded', function() {
            console.log('打印页面数量:', document.querySelectorAll('.print-page').length);
            console.log('字段元素数量:', document.querySelectorAll('.field-element').length);
            
            // 确保所有元素可见
            document.querySelectorAll('.field-element').forEach(function(el) {
                el.style.display = 'block';
                el.style.visibility = 'visible';
                el.style.overflow = 'visible';
            });
            
            setTimeout(function() {
                console.log('准备打印...');
                window.print();
            }, 1000);
        });
    </script>
</body>
</html>`;

        // 合并完整的HTML内容
        const printHTML = htmlHead + htmlBody + htmlFoot;
        
        // 直接使用内嵌iframe模式显示打印预览
        console.log('使用内嵌iframe模式显示打印预览');
        const printModal = createInlinePreviewModal(printHTML, '打印预览');
        
        // 添加直接打印按钮
        const printButton = document.createElement('button');
        printButton.textContent = '打印';
        printButton.style.position = 'absolute';
        printButton.style.right = '80px';
        printButton.style.top = '15px';
        printButton.style.zIndex = '10000';
        printButton.style.padding = '5px 15px';
        printButton.style.backgroundColor = '#4a90e2';
        printButton.style.color = 'white';
        printButton.style.border = 'none';
        printButton.style.borderRadius = '4px';
        printButton.style.cursor = 'pointer';
        
        printButton.addEventListener('click', function() {
            const iframe = printModal.iframe;
            if (iframe && iframe.contentWindow) {
                iframe.contentWindow.print();
            }
        });
        
        // 添加按钮到模态框
        printModal.modal.querySelector('.modal-content').appendChild(printButton);
    }
    
    // 清空画布函数
    function clearCanvas() {
        // 保存当前画布尺寸
        const currentWidth = canvas.style.width;
        const currentHeight = canvas.style.height;
        
        // 清除所有子元素
        while (canvas.firstChild) {
            canvas.removeChild(canvas.firstChild);
        }
        
        // 重置变量
        templateImage = null;
        fields = [];
        selectedField = null;
        
        // 添加占位信息
        const placeholder = document.createElement('div');
        placeholder.className = 'placeholder-message';
        placeholder.textContent = '请导入模板或创建新模板';
        canvas.appendChild(placeholder);
        
        // 恢复画布尺寸
        canvas.style.width = currentWidth;
        canvas.style.height = currentHeight;
    }
    
    // 开始拖动字段
    function handleFieldMouseDown(e) {
        if (e.target.classList.contains('resize-handle') || isResizing) {
            return; // 如果点击的是调整大小的控制点或正在调整大小，不启动拖动
        }
        
        if (e.target.classList.contains('field')) {
            e.stopPropagation();
            
            // 选中字段
            selectField(e.target);
            
            // 开始拖动
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            
            const styles = window.getComputedStyle(e.target);
            startLeft = parseInt(styles.left, 10);
            startTop = parseInt(styles.top, 10);
            
            // 创建辅助线容器
            removeAllAlignmentGuides();
            
            // 添加鼠标移动事件
            document.addEventListener('mousemove', handleFieldMouseMove);
            document.addEventListener('mouseup', handleFieldMouseUp);
        }
    }
    
    function handleFieldMouseMove(e) {
        if (isDragging && selectedField && !isResizing) {
            // 计算鼠标移动的距离，考虑缩放因素
            const dx = (e.clientX - startX) / currentZoom;
            const dy = (e.clientY - startY) / currentZoom;
            
            // 修复：直接使用startLeft和startTop，它们已经是正确的显示像素位置
            // 计算新的显示位置
            let newLeft = startLeft + dx;
            let newTop = startTop + dy;
            
            // 移除之前的辅助线
            removeAllAlignmentGuides();
            
            // 获取当前字段的尺寸
            const fieldWidth = parseInt(selectedField.style.width, 10);
            const fieldHeight = parseInt(selectedField.style.height, 10);
            
            // 计算当前字段的中心点和边界点（像素）
            const selectedCenterX = newLeft + fieldWidth / 2;
            const selectedCenterY = newTop + fieldHeight / 2;
            const selectedRightX = newLeft + fieldWidth;
            const selectedBottomY = newTop + fieldHeight;
            
            // 用于存储找到的对齐点
            let alignX = null;
            let alignY = null;
            
            // 吸附阈值（毫米）
            const snapThreshold = 2;
            // 计算吸附阈值（像素）
            const snapThresholdPx = mmToPx(snapThreshold);
            
            // 遍历其他字段，检查对齐
            fields.forEach(field => {
                // 跳过当前选中的字段
                if (field.id === selectedField.id) return;
                
                // 获取其他字段的元素
                const otherField = document.getElementById(field.id);
                if (!otherField) return;
                
                // 获取其他字段的位置和尺寸（像素）
                const otherX = parseInt(otherField.style.left, 10);
                const otherY = parseInt(otherField.style.top, 10);
                const otherWidth = parseInt(otherField.style.width, 10);
                const otherHeight = parseInt(otherField.style.height, 10);
                
                // 计算其他字段的中心点和边界点
                const otherCenterX = otherX + otherWidth / 2;
                const otherCenterY = otherY + otherHeight / 2;
                const otherRightX = otherX + otherWidth;
                const otherBottomY = otherY + otherHeight;
                
                // 检查左边缘对齐
                if (Math.abs(newLeft - otherX) < snapThresholdPx) {
                    alignX = otherX;
                    createAlignmentGuide('vertical', alignX);
                }
                // 检查右边缘对齐
                else if (Math.abs(selectedRightX - otherRightX) < snapThresholdPx) {
                    alignX = otherRightX - fieldWidth;
                    createAlignmentGuide('vertical', otherRightX);
                }
                // 检查中心对齐
                else if (Math.abs(selectedCenterX - otherCenterX) < snapThresholdPx) {
                    alignX = otherCenterX - fieldWidth / 2;
                    createAlignmentGuide('vertical', otherCenterX);
                }
                // 检查左边缘与右边缘对齐
                else if (Math.abs(selectedRightX - otherX) < snapThresholdPx) {
                    alignX = otherX - fieldWidth;
                    createAlignmentGuide('vertical', otherX);
                }
                // 检查右边缘与左边缘对齐
                else if (Math.abs(newLeft - otherRightX) < snapThresholdPx) {
                    alignX = otherRightX;
                    createAlignmentGuide('vertical', otherRightX);
                }
                
                // 检查顶部对齐
                if (Math.abs(newTop - otherY) < snapThresholdPx) {
                    alignY = otherY;
                    createAlignmentGuide('horizontal', alignY);
                }
                // 检查底部对齐
                else if (Math.abs(selectedBottomY - otherBottomY) < snapThresholdPx) {
                    alignY = otherBottomY - fieldHeight;
                    createAlignmentGuide('horizontal', otherBottomY);
                }
                // 检查中心对齐
                else if (Math.abs(selectedCenterY - otherCenterY) < snapThresholdPx) {
                    alignY = otherCenterY - fieldHeight / 2;
                    createAlignmentGuide('horizontal', otherCenterY);
                }
                // 检查顶部与底部对齐
                else if (Math.abs(selectedBottomY - otherY) < snapThresholdPx) {
                    alignY = otherY - fieldHeight;
                    createAlignmentGuide('horizontal', otherY);
                }
                // 检查底部与顶部对齐
                else if (Math.abs(newTop - otherBottomY) < snapThresholdPx) {
                    alignY = otherBottomY;
                    createAlignmentGuide('horizontal', otherBottomY);
                }
            });
            
            // 应用吸附位置（如果找到）
            if (alignX !== null) {
                newLeft = alignX;
            }
            if (alignY !== null) {
                newTop = alignY;
            }
            
            // 更新字段DOM位置
            selectedField.style.left = newLeft + 'px';
            selectedField.style.top = newTop + 'px';
            
            // 修复：直接将显示位置转换为毫米，不要再除以currentZoom
            const mmX = pxToMm(newLeft);
            const mmY = pxToMm(newTop);
            
            // 更新属性面板
            document.getElementById('field-x').value = parseFloat(mmX.toFixed(1));
            document.getElementById('field-y').value = parseFloat(mmY.toFixed(1));
            
            // 更新字段对象，确保保存的是正确的毫米位置
            const field = fields.find(f => f.id === selectedField.id);
            if (field) {
                field.x = parseFloat(mmX.toFixed(1));
                field.y = parseFloat(mmY.toFixed(1));
            }
        }
    }
    
    function handleFieldMouseUp(e) {
        if (isDragging) {
            isDragging = false;
            
            // 清除所有对齐辅助线
            removeAllAlignmentGuides();
            
            // 添加自动保存
            saveToLocalStorage();
            
            // 移除鼠标事件监听器
            document.removeEventListener('mousemove', handleFieldMouseMove);
            document.removeEventListener('mouseup', handleFieldMouseUp);
        }
    }
    
    // 放大模板
    function zoomIn() {
        currentZoom += 0.1;
        applyZoom();
    }
    
    // 缩小模板
    function zoomOut() {
        // 限制最小缩放比例为0.5
        if (currentZoom > 0.5) {
            currentZoom -= 0.1;
            applyZoom();
        }
    }
    
    // 应用缩放
    function applyZoom() {
        // 更新画布缩放比例
        canvas.style.transform = `scale(${currentZoom})`;
        canvas.style.transformOrigin = 'top left';
        
        // 更新画布容器的高度以适应缩放后的内容
        const templateArea = document.querySelector('.template-area');
        const canvasHeight = canvas.offsetHeight * currentZoom;
        templateArea.style.minHeight = `${canvasHeight}px`;
        
        // 更新所有字段的位置，确保在缩放变化时保持正确位置
        fields.forEach(field => {
            const fieldElement = document.getElementById(field.id);
            if (fieldElement) {
                // 使用毫米位置重新计算像素位置
                fieldElement.style.left = mmToPx(field.x) + 'px';
                fieldElement.style.top = mmToPx(field.y) + 'px';
            }
        });
        
        // 计算浏览器自动缩放比例（毫米到像素的转换）
        const mmToPixelRatio = canvas.offsetWidth / templateWidth;
        
        // 计算总的实际缩放比例
        const totalZoom = mmToPixelRatio * currentZoom;
        
        console.log("手动缩放比例:", currentZoom);
        console.log("浏览器自动缩放比例:", mmToPixelRatio.toFixed(3));
        console.log("总实际缩放比例:", totalZoom.toFixed(3));
    }
    
    // 确认模板尺寸函数
    function confirmTemplateSize() {
        const newWidth = parseInt(templateWidthInput.value);
        const newHeight = parseInt(templateHeightInput.value);
        
        if (newWidth > 0 && newHeight > 0) {
            // 保存原始尺寸用于比例计算
            const oldWidth = templateWidth;
            const oldHeight = templateHeight;
            
            // 更新模板尺寸
            templateWidth = newWidth;
            templateHeight = newHeight;
            
            // 应用新尺寸到画布
            canvas.style.width = templateWidth + 'mm';
            canvas.style.height = templateHeight + 'mm';
            
            // 如果已有字段，根据比例调整字段位置
            if (fields.length > 0) {
                const widthRatio = newWidth / oldWidth;
                const heightRatio = newHeight / oldHeight;
                
                fields.forEach(field => {
                    // 调整字段位置和宽度
                    field.x = Math.round(field.x * widthRatio);
                    field.y = Math.round(field.y * heightRatio);
                    field.width = Math.round(field.width * widthRatio);
                    
                    // 更新DOM元素
                    const fieldElement = document.getElementById(field.id);
                    if (fieldElement) {
                        fieldElement.style.left = field.x + 'px';
                        fieldElement.style.top = field.y + 'px';
                        fieldElement.style.width = field.width + 'px';
                    }
                });
            }
            
            // 隐藏模态框
            templateSizeModal.style.display = 'none';
        } else {
            alert('请输入有效的尺寸数值！');
        }
        

        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 导出数据函数
    function exportData() {
        // 检查是否有模板字段
        if (fields.length === 0) {
            alert('没有可导出的字段数据！请先添加字段。');
            return;
        }
        
        try {
            // 创建工作簿
            const wb = XLSX.utils.book_new();
            
            // 准备数据
            const excelData = [];
            
            // 添加表头行
            const headers = fields.map(field => field.name);
            excelData.push(headers);
            
            // 始终只导出一行默认值，不管是否有导入的数据
            const defaultValues = fields.map(field => {
                let fieldValue;
                
                if (field.type === 'image') {
                    // 对于图片字段，导出图片内容（base64数据）而不是路径
                    fieldValue = field.imageDataUrl || '';
                } else {
                    fieldValue = field.defaultValue || '';
                }
                
                return fieldValue;
            });
            
            excelData.push(defaultValues);
            
            // 创建工作表
            const ws = XLSX.utils.aoa_to_sheet(excelData);
            
            // 将工作表添加到工作簿
            XLSX.utils.book_append_sheet(wb, ws, "数据");
            
            // 设置文件名
            const now = new Date();
            const timestamp = now.getFullYear() + 
                          ("0" + (now.getMonth() + 1)).slice(-2) + 
                          ("0" + now.getDate()).slice(-2) + 
                          ("0" + now.getHours()).slice(-2) + 
                          ("0" + now.getMinutes()).slice(-2);
            
            const fileName = "套打数据_" + timestamp + ".xlsx";
            
            // 导出Excel文件
            XLSX.writeFile(wb, fileName);
            
            console.log("数据导出成功:", fileName);
            
        } catch (e) {
            console.error("导出数据时出错:", e);
            alert("导出数据失败，请重试！");
        }
    }
    
    // 移除对齐辅助线函数
    function removeAlignmentGuide(type) {
        if (alignmentGuides[type]) {
            alignmentGuides[type].remove();
            alignmentGuides[type] = null;
        }
    }
    
    // 移除所有对齐辅助线
    function removeAllAlignmentGuides() {
        removeAlignmentGuide('vertical');
        removeAlignmentGuide('horizontal');
    }
    
    // 初始化时添加字体选择下拉菜单的事件监听
    document.addEventListener('DOMContentLoaded', function() {
        // 字体选择下拉菜单变化监听
        const fontFamilySelect = document.getElementById('field-font-family');
        const customFontContainer = document.getElementById('custom-font-container');
        
        fontFamilySelect.addEventListener('change', function() {
            if (this.value === 'custom') {
                customFontContainer.style.display = 'block';
                document.getElementById('field-custom-font').focus();
            } else {
                customFontContainer.style.display = 'none';
            }
        });
        
        // 字体样式按钮点击事件
        document.querySelectorAll('.style-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                this.classList.toggle('active');
                if (selectedField) {
                    applyFontStyles();
                } else {
                    alert('请先选中一个字段');
                }
            });
        });
    });
    
    // 处理开始调整大小事件
    function handleResizeStart(e, fieldElement, position) {
        // 阻止事件传播，避免触发字段的拖动
        e.stopPropagation();
        e.preventDefault();
        console.log("开始调整大小", position);
        
        // 确保字段被选中
        if (fieldElement !== selectedField) {
            selectField(fieldElement);
        }
        
        // 获取初始位置和大小
        const styles = window.getComputedStyle(fieldElement);
        const startLeft = parseInt(fieldElement.style.left, 10) || 0;
        const startTop = parseInt(fieldElement.style.top, 10) || 0;
        const startWidth = parseInt(styles.width, 10);
        const startHeight = parseInt(styles.height, 10);
        const startX = e.clientX;
        const startY = e.clientY;
        
        // 获取字段对象
        const field = fields.find(f => f.id === fieldElement.id);
        if (!field) return;
        
        // 禁用画布的拖动
        document.body.style.userSelect = 'none';
        isResizing = true;
        
        // 添加鼠标移动和抬起事件监听
        function handleResizeMove(moveEvent) {
            moveEvent.preventDefault();
            
            // 计算鼠标移动的距离，考虑缩放因素
            const dx = (moveEvent.clientX - startX) / currentZoom;
            const dy = (moveEvent.clientY - startY) / currentZoom;
            
            // 根据不同的控制点位置，更新尺寸和位置
            let newWidth = startWidth;
            let newHeight = startHeight;
            let newLeft = startLeft;
            let newTop = startTop;
            
            if (position.includes('e')) {
                newWidth = Math.max(30, startWidth + dx);
            }
            if (position.includes('w')) {
                const widthChange = Math.min(startWidth - 30, dx);
                newWidth = startWidth - widthChange;
                newLeft = startLeft + widthChange;
            }
            if (position.includes('s')) {
                newHeight = Math.max(20, startHeight + dy);
            }
            if (position.includes('n')) {
                const heightChange = Math.min(startHeight - 20, dy);
                newHeight = startHeight - heightChange;
                newTop = startTop + heightChange;
            }
            
            // 更新DOM元素
            fieldElement.style.width = newWidth + 'px';
            fieldElement.style.height = newHeight + 'px';
            fieldElement.style.left = newLeft + 'px';
            fieldElement.style.top = newTop + 'px';
            
            // 将像素值转换为毫米
            const mmWidth = pxToMm(newWidth);
            const mmHeight = pxToMm(newHeight);
            const mmX = pxToMm(newLeft);
            const mmY = pxToMm(newTop);
            
            // 更新字段对象
            field.width = parseFloat(mmWidth.toFixed(1));
            field.height = parseFloat(mmHeight.toFixed(1));
            field.x = parseFloat(mmX.toFixed(1));
            field.y = parseFloat(mmY.toFixed(1));
            
            // 更新属性面板
            document.getElementById('field-width').value = field.width;
            document.getElementById('field-height').value = field.height;
            document.getElementById('field-x').value = field.x;
            document.getElementById('field-y').value = field.y;
            
            console.log("调整大小中", newWidth, newHeight, newLeft, newTop);
        }
        
        function handleResizeEnd() {
            document.removeEventListener('mousemove', handleResizeMove);
            document.removeEventListener('mouseup', handleResizeEnd);
            
            // 恢复用户选择和拖动状态
            document.body.style.userSelect = '';
            isResizing = false;
            console.log("调整大小结束");
        }
        
        document.addEventListener('mousemove', handleResizeMove);
        document.addEventListener('mouseup', handleResizeEnd);
    }
    

    
    // 处理字段双击事件
    function handleFieldDoubleClick(fieldElement) {
        // 获取字段对象
        const field = fields.find(f => f.id === fieldElement.id);
        if (!field) return;
        
        // 获取当前内容
        let currentContent = '';
        
        // 根据字段类型处理
        if (field.type === 'image') {
            // 对于图片字段，使用空白输入框，不显示原有Base64数据
            currentContent = ''; // 空白输入框
        } else {
            // 对于其他字段，使用默认值或名称
            currentContent = fieldElement.childNodes[0].nodeValue || field.defaultValue || field.name;
        }
        
        // 创建输入框
        const inputElement = document.createElement('textarea'); // 使用textarea以便支持长文本
        inputElement.value = currentContent;
        inputElement.style.width = '100%';
        inputElement.style.height = '100%';
        inputElement.style.border = 'none';
        inputElement.style.padding = '0';
        inputElement.style.fontSize = 'inherit';
        inputElement.style.fontFamily = 'inherit';
        inputElement.style.fontWeight = 'inherit';
        inputElement.style.background = 'transparent';
        inputElement.style.outline = 'none';
        inputElement.style.resize = 'none';
        
        // 为图片字段添加占位符提示
        if (field.type === 'image') {
            inputElement.placeholder = '在此粘贴Base64图片数据';
        }
        
        // 临时保存字段标签
        const fieldLabel = fieldElement.querySelector('.field-label');
        
        // 清空字段内容
        fieldElement.textContent = '';
        
        // 添加输入框和标签
        fieldElement.appendChild(inputElement);
        if (fieldLabel) fieldElement.appendChild(fieldLabel);
        
        // 聚焦输入框
        inputElement.focus();
        
        // 处理输入框失焦事件
        inputElement.addEventListener('blur', function() {
            saveFieldContent(fieldElement, inputElement.value, field);
        });
        
        // 处理回车键事件
        inputElement.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                inputElement.blur(); // 触发失焦事件
            }
        });
    }
    
    // 保存字段内容
    function saveFieldContent(fieldElement, content, field) {
        // 临时保存字段标签
        const fieldLabel = fieldElement.querySelector('.field-label');
        
        // 清空字段内容
        fieldElement.textContent = '';
        
        // 根据字段类型处理内容
        if (field.type === 'image') {
            // 检查内容是否为有效的Base64图片数据
            if (content.trim() === '') {
                // 如果输入框为空，保持原有图片不变
                if (field.imageDataUrl) {
                    // 创建图片元素
                    const img = document.createElement('img');
                    img.src = field.imageDataUrl;
                    img.alt = field.name;
                    img.style.maxWidth = '100%';
                    img.style.maxHeight = '100%';
                    
                    // 添加到字段元素
                    fieldElement.appendChild(img);
                } else {
                    // 没有原有图片，显示提示
                    fieldElement.textContent = '无图片';
                }
            } else if (content.startsWith('data:image/')) {
                // 更新字段对象
                field.imageDataUrl = content;
                
                // 创建图片元素
                const img = document.createElement('img');
                img.src = content;
                img.alt = field.name;
                img.style.maxWidth = '100%';
                img.style.maxHeight = '100%';
                
                // 添加到字段元素
                fieldElement.appendChild(img);
            } else {
                // 不是有效的Base64数据，显示提示
                fieldElement.textContent = '无效的图片数据';
            }
        } else {
            // 更新字段对象
            field.defaultValue = content;
            
            // 更新DOM元素内容
            const contentSpan = document.createElement('span');
            contentSpan.className = 'field-content';
            contentSpan.textContent = content;
            contentSpan.style.fontWeight = field.fontWeight || 'normal';
            contentSpan.style.fontStyle = field.fontStyle || 'normal';
            contentSpan.style.textDecoration = field.textDecoration || 'none';
            fieldElement.appendChild(contentSpan);
        }
        
        // 重新添加字段标签
        if (fieldLabel) fieldElement.appendChild(fieldLabel);
        
        // 保存到本地存储
        saveToLocalStorage();
    }
    
    // 创建对齐辅助线函数
    function createAlignmentGuide(type, position) {
        // 移除现有的辅助线
        removeAlignmentGuide(type);
        
        // 创建新的辅助线
        const guide = document.createElement('div');
        guide.className = `alignment-guide ${type}`;
        
        if (type === 'vertical') {
            guide.style.left = position + 'px';
            guide.style.height = canvas.offsetHeight + 'px';
        } else {
            guide.style.top = position + 'px';
            guide.style.width = canvas.offsetWidth + 'px';
        }
        
        // 添加到画布
        canvas.appendChild(guide);
        
        // 保存辅助线引用
        alignmentGuides[type] = guide;
    }

    // 初始化IndexedDB
    function initIndexedDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            
            request.onerror = (event) => {
                console.error('IndexedDB打开失败:', event.target.error);
                reject(event.target.error);
            };
            
            request.onsuccess = (event) => {
                db = event.target.result;
                console.log('IndexedDB连接成功');
                resolve(db);
            };
            
            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'id' });
                    console.log('创建数据存储对象成功');
                }
            };
        });
    }

    // 保存到IndexedDB
    function saveToIndexedDB() {
        if (!db) {
            console.error('IndexedDB未初始化');
            return Promise.reject('IndexedDB未初始化');
        }
        
        const dataToSave = {
            id: TEMPLATE_KEY,
            templateWidth: templateWidth,
            templateHeight: templateHeight,
            fields: fields,
            templateImage: templateImage ? templateImage.src : null, // 保存完整图片数据
            templateData: templateData,
            currentZoom: currentZoom,
            printOffset: window.savedPrintOffset || { x: 0, y: 0 },
            timestamp: new Date().getTime()
        };
        
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readwrite');
            const store = transaction.objectStore(STORE_NAME);
            
            const request = store.put(dataToSave); // 使用put而不是add，以便覆盖旧数据
            
            request.onsuccess = () => {
                console.log('数据已保存到IndexedDB');
                resolve();
            };
            
            request.onerror = (event) => {
                console.error('保存到IndexedDB失败:', event.target.error);
                reject(event.target.error);
            };
        });
    }

    // 保存到本地存储
    function saveToLocalStorage() {
        // 尝试使用IndexedDB保存
        saveToIndexedDB().catch(error => {
            console.error('IndexedDB保存失败，尝试使用localStorage:', error);
            
            // 如果IndexedDB失败，回退到localStorage，但不保存图片数据
            const dataToSave = {
                templateWidth: templateWidth,
                templateHeight: templateHeight,
                fields: fields,
                hasTemplateImage: templateImage ? true : false, // 只保存是否有图片的标志
                templateData: templateData,
                currentZoom: currentZoom,
                printOffset: window.savedPrintOffset || { x: 0, y: 0 }
            };
            
            try {
                localStorage.setItem(TEMPLATE_KEY, JSON.stringify(dataToSave));
                console.log('数据已保存到本地存储（不包含图片）');
            } catch (e) {
                console.error('保存到本地存储失败:', e);
            }
        });
    }

    // 从IndexedDB加载
    function loadFromIndexedDB() {
        if (!db) {
            console.error('IndexedDB未初始化');
            return Promise.reject('IndexedDB未初始化');
        }
        
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            
            const request = store.get(TEMPLATE_KEY);
            
            request.onsuccess = (event) => {
                const data = event.target.result;
                if (data) {
                    console.log('从IndexedDB加载数据成功');
                    resolve(data);
                } else {
                    console.log('IndexedDB中没有找到数据');
                    resolve(null);
                }
            };
            
            request.onerror = (event) => {
                console.error('从IndexedDB加载失败:', event.target.error);
                reject(event.target.error);
            };
        });
    }

    // 从存储加载
    function loadFromLocalStorage() {
        // 首先尝试从IndexedDB加载
        if (db) {
            loadFromIndexedDB().then(data => {
                if (data) {
                    // 使用IndexedDB中的数据
                    applyLoadedData(data);
                } else {
                    // 如果IndexedDB中没有数据，尝试从localStorage加载
                    console.log('IndexedDB中没有找到数据，尝试从localStorage加载');
                    loadFromLocalStorageFallback();
                }
            }).catch(error => {
                console.error('从IndexedDB加载失败，尝试使用localStorage:', error);
                loadFromLocalStorageFallback();
            });
        } else {
            console.log('IndexedDB未初始化，从localStorage加载');
            loadFromLocalStorageFallback();
        }
    }
    
    // 从localStorage备份加载
    function loadFromLocalStorageFallback() {
        try {
            const savedData = localStorage.getItem(TEMPLATE_KEY);
            if (savedData) {
                const data = JSON.parse(savedData);
                console.log('从localStorage加载数据成功');
                applyLoadedData(data);
            }
        } catch (e) {
            console.error('从localStorage加载数据失败:', e);
        }
    }
    
    // 应用加载的数据
    function applyLoadedData(data) {
        // 恢复模板尺寸
        templateWidth = data.templateWidth || 210;
        templateHeight = data.templateHeight || 297;
        canvas.style.width = templateWidth + 'mm';
        canvas.style.height = templateHeight + 'mm';
        
        // 恢复缩放比例
        currentZoom = data.currentZoom || 1;
        applyZoom();
        
        // 恢复背景图片
        if (data.templateImage) {
            const img = new Image();
            img.src = data.templateImage;
            img.onload = function() {
                templateImage = document.createElement('img');
                templateImage.src = data.templateImage;
                templateImage.className = 'template-image';
                templateImage.style.width = '100%';
                templateImage.style.height = '100%';
                templateImage.style.position = 'absolute';
                templateImage.style.top = '0';
                templateImage.style.left = '0';
                templateImage.style.zIndex = '0';
                canvas.appendChild(templateImage);
                
                // 移除占位消息
                const placeholder = canvas.querySelector('.placeholder-message');
                if (placeholder) {
                    placeholder.style.display = 'none';
                }
            };
            
            img.onerror = function() {
                console.warn('保存的背景图片加载失败');
                // 移除占位消息
                const placeholder = canvas.querySelector('.placeholder-message');
                if (placeholder) {
                    placeholder.style.display = 'none';
                }
            };
        }
        
        // 恢复字段
        fields = data.fields || [];
        fields.forEach(field => {
            createFieldElement(field);
        });
        
        // 恢复模板数据
        templateData = data.templateData || [];
        
        // 恢复打印偏移值
        window.savedPrintOffset = data.printOffset || { x: 0, y: 0 };
    }

    // 增强属性面板交互
    function enhancePropertiesPanel() {
        // 为属性分组添加折叠功能
        const sectionTitles = document.querySelectorAll('.properties-section-title');
        
        sectionTitles.forEach(title => {
            // 添加折叠图标
            const icon = document.createElement('span');
            icon.innerHTML = '▼';
            icon.style.fontSize = '12px';
            icon.style.marginRight = '5px';
            icon.style.transition = 'transform 0.3s';
            title.prepend(icon);
            
            // 添加点击事件
            title.style.cursor = 'pointer';
            title.addEventListener('click', () => {
                const section = title.closest('.properties-section');
                const content = section.querySelectorAll('.property-group, .coords-group, .font-style-options');
                
                // 切换显示/隐藏
                let isVisible = section.dataset.expanded !== 'false';
                isVisible = !isVisible;
                
                content.forEach(el => {
                    el.style.display = isVisible ? 'flex' : 'none';
                });
                
                // 更新图标
                icon.style.transform = isVisible ? 'rotate(0deg)' : 'rotate(-90deg)';
                section.dataset.expanded = isVisible;
            });
        });
        
        // 为输入框添加动画效果
        const inputs = document.querySelectorAll('.property-group input, .property-group select');
        inputs.forEach(input => {
            input.addEventListener('focus', () => {
                input.parentElement.style.transform = 'translateX(5px)';
                input.parentElement.style.transition = 'transform 0.2s';
            });
            
            input.addEventListener('blur', () => {
                input.parentElement.style.transform = 'translateX(0)';
            });
        });
    }
    
    // 常用尺寸（横版，单位mm）
    const sizeMap = {
        'A3': { width: 420, height: 297 },
        'A4': { width: 297, height: 210 },
        'A5': { width: 210, height: 148 },
        'B5': { width: 250, height: 176 }
    };

    if (templateSizeSelect) {
        templateSizeSelect.addEventListener('change', function() {
            const val = this.value;
            if (sizeMap[val]) {
                // 使用标准尺寸的宽度和高度
                templateWidthInput.value = sizeMap[val].width;
                templateHeightInput.value = sizeMap[val].height;
                
                // 将宽度和高度都设为只读
                templateWidthInput.readOnly = true;
                templateHeightInput.readOnly = true;
            } else {
                templateWidthInput.readOnly = false;
                templateHeightInput.readOnly = false;
            }
        });
        // 初始化时根据默认选项设置
        if (sizeMap[templateSizeSelect.value]) {
            templateWidthInput.value = sizeMap[templateSizeSelect.value].width;
            templateHeightInput.value = sizeMap[templateSizeSelect.value].height;
            templateWidthInput.readOnly = true;
            templateHeightInput.readOnly = true;
        }
    }

    // 创建打印HTML的函数
    function createPrintHTML(dataToPrint, shouldShowBackground, xOffset, yOffset) {
        // 创建HTML头部内容
        const htmlHead = `<!DOCTYPE html>
<html>
<head>
    <title>套打打印</title>
    <meta charset="UTF-8">
    <style>
        @page {
            size: ${templateWidth}mm ${templateHeight}mm;
            margin: 0;
        }
        body {
            margin: 0;
            padding: 0;
            overflow: hidden;
        }
        .print-page {
            position: relative;
            width: ${templateWidth}mm;
            height: ${templateHeight}mm;
            margin: 0;
            padding: 0;
            page-break-after: always;
            overflow: visible;
        }
        .print-page:last-child {
            page-break-after: avoid;
        }
        .field-element {
            position: absolute;
            overflow: visible;
            white-space: normal;
            z-index: 10;
        }
        .background-image {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 0;
        }
        @media print {
            body {
                margin: 0;
                padding: 0;
            }
            .print-page {
                display: block !important;
                visibility: visible !important;
                overflow: visible !important;
            }
            .field-element {
                display: block !important;
                visibility: visible !important;
                overflow: visible !important;
            }
        }
    </style>
</head>
<body>`;

        // 创建HTML页面内容
        let htmlBody = '';
        
        // 为每组数据创建一个打印页面
        dataToPrint.forEach((data, index) => {
            htmlBody += '<div class="print-page">';
            
            // 添加背景图片（如果选择显示）
            if (shouldShowBackground && templateImage) {
                htmlBody += `<img class="background-image" src="${templateImage.src}" alt="模板背景" />`;
            }
            
            // 添加所有字段
        fields.forEach(field => {
                // 尝试多种可能的字段名匹配方式
                let fieldValue = '';
                
                // 1. 直接匹配字段名
                if (data[field.name] !== undefined) {
                    fieldValue = data[field.name];
                }
                // 2. 尝试去除引号和空格后匹配
                else if (data[field.name.trim()] !== undefined) {
                    fieldValue = data[field.name.trim()];
                }
                // 3. 尝试转为小写后匹配
                else if (data[field.name.toLowerCase()] !== undefined) {
                    fieldValue = data[field.name.toLowerCase()];
                }
                // 4. 尝试转为大写后匹配
                else if (data[field.name.toUpperCase()] !== undefined) {
                    fieldValue = data[field.name.toUpperCase()];
                }
                // 5. 如果都没匹配到，使用默认值
                else {
                    fieldValue = field.defaultValue || '';
                }
                
                // 处理图片字段
                if (field.type === 'image' && fieldValue) {
                    // 如果是BASE64数据，创建img元素
                    if (fieldValue.startsWith('data:image/') || fieldValue.length > 100) {
                        htmlBody += `<div class="field-element" 
                                style="left: ${field.x + xOffset}mm; 
                                       top: ${field.y + yOffset}mm; 
                                       width: ${field.width}mm; 
                                       height: ${field.height}mm;">
                               <img src="${fieldValue}" alt="${field.name}" style="width: 100%; height: 100%; object-fit: contain;" />
                           </div>`;
                    } else {
                        // 如果不是BASE64，可能是文件路径，显示为普通文本
                        htmlBody += `<div class="field-element" 
                                style="left: ${field.x + xOffset}mm; 
                                       top: ${field.y + yOffset}mm; 
                                       width: ${field.width}mm; 
                                       height: ${field.height}mm;
                                       font-size: ${field.fontSize}pt; 
                                       font-family: ${field.fontFamily};
                                       font-weight: ${field.fontWeight || 'normal'};
                                       font-style: ${field.fontStyle || 'normal'};
                                       text-decoration: ${field.textDecoration || 'none'};">
                               ${fieldValue}
                           </div>`;
                    }
                } else {
                    // 普通字段处理
                    htmlBody += `<div class="field-element" 
                            style="left: ${field.x + xOffset}mm; 
                                   top: ${field.y + yOffset}mm; 
                                   width: ${field.width}mm; 
                                   height: ${field.height}mm;
                                   font-size: ${field.fontSize}pt; 
                                   font-family: ${field.fontFamily};
                                   font-weight: ${field.fontWeight || 'normal'};
                                   font-style: ${field.fontStyle || 'normal'};
                                   text-decoration: ${field.textDecoration || 'none'};">
                           ${fieldValue}
                       </div>`;
                }
            });
            
            htmlBody += '</div>';
        });
        
        // 创建HTML尾部和脚本
        const htmlFoot = `</body>\n</html>`;

        // 返回完整的HTML内容
        return htmlHead + htmlBody + htmlFoot;
    }

    document.getElementById('field-font-bold').addEventListener('change', function() {
        if (selectedField) applyFontStyles();
    });
    document.getElementById('field-font-italic').addEventListener('change', function() {
        if (selectedField) applyFontStyles();
    });
    document.getElementById('field-font-underline').addEventListener('change', function() {
        if (selectedField) applyFontStyles();
    });
});