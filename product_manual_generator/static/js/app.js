/**
 * 产品手册生成器 - 前端交互脚本
 * 处理文件上传、配置管理和生成请求
 */

// ===== 全局状态 =====
const state = {
    templateUploaded: false,
    imagesUploaded: false,
    priceUploaded: false,
    uploadedImages: [],
    templateInfo: null,
    priceInfo: null
};

// ===== 初始化 =====
document.addEventListener('DOMContentLoaded', function() {
    initUploadAreas();
    initTabs();
    initDragAndDrop();
});

// ===== 上传区域初始化 =====
function initUploadAreas() {
    // PPT模板上传
    const templateArea = document.getElementById('template-upload-area');
    const templateInput = document.getElementById('template-input');
    
    templateArea.addEventListener('click', () => templateInput.click());
    templateInput.addEventListener('change', (e) => handleTemplateUpload(e.target.files[0]));

    // 图片上传
    const imagesArea = document.getElementById('images-upload-area');
    const imagesInput = document.getElementById('images-input');
    
    imagesArea.addEventListener('click', () => imagesInput.click());
    imagesInput.addEventListener('change', (e) => handleImagesUpload(e.target.files));

    // ZIP上传
    const zipArea = document.getElementById('zip-upload-area');
    const zipInput = document.getElementById('zip-input');
    
    zipArea.addEventListener('click', () => zipInput.click());
    zipInput.addEventListener('change', (e) => handleZipUpload(e.target.files[0]));

    // 价格表上传
    const priceArea = document.getElementById('price-upload-area');
    const priceInput = document.getElementById('price-input');
    
    priceArea.addEventListener('click', () => priceInput.click());
    priceInput.addEventListener('change', (e) => handlePriceUpload(e.target.files[0]));
}

// ===== 标签页切换 =====
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;
            
            // 切换按钮状态
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // 切换内容
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(`tab-${tab}`).classList.add('active');
        });
    });
}

// ===== 拖拽上传初始化 =====
function initDragAndDrop() {
    const uploadAreas = document.querySelectorAll('.upload-area');
    
    uploadAreas.forEach(area => {
        area.addEventListener('dragover', (e) => {
            e.preventDefault();
            area.classList.add('dragover');
        });
        
        area.addEventListener('dragleave', () => {
            area.classList.remove('dragover');
        });
        
        area.addEventListener('drop', (e) => {
            e.preventDefault();
            area.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                // 根据区域类型处理文件
                if (area.id === 'template-upload-area') {
                    handleTemplateUpload(files[0]);
                } else if (area.id === 'images-upload-area') {
                    handleImagesUpload(files);
                } else if (area.id === 'zip-upload-area') {
                    handleZipUpload(files[0]);
                } else if (area.id === 'price-upload-area') {
                    handlePriceUpload(files[0]);
                }
            }
        });
    });
}

// ===== PPT模板上传处理 =====
async function handleTemplateUpload(file) {
    if (!file) return;
    
    if (!file.name.endsWith('.pptx')) {
        addLog('请上传 .pptx 格式的PPT文件', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('template', file);
    
    try {
        addLog('正在上传PPT模板...', 'info');
        
        const response = await fetch('/api/upload/template', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            state.templateUploaded = true;
            state.templateInfo = data.template_info;
            
            // 更新UI
            document.getElementById('template-info').style.display = 'flex';
            document.getElementById('template-info').querySelector('.file-name').textContent = file.name;
            document.getElementById('template-upload-area').style.display = 'none';
            document.getElementById('template-status').classList.add('success');
            
            // 显示模板信息
            document.getElementById('template-preview').style.display = 'block';
            document.getElementById('preview-slides').textContent = data.template_info.slide_count;
            document.getElementById('preview-positions').textContent = data.template_info.product_positions;
            
            addLog(`PPT模板上传成功！共 ${data.template_info.slide_count} 页，${data.template_info.product_positions} 个产品位置`, 'success');
        } else {
            addLog(`上传失败：${data.error}`, 'error');
        }
    } catch (error) {
        addLog(`上传出错：${error.message}`, 'error');
    }
}

// ===== 移除模板 =====
function removeTemplate() {
    state.templateUploaded = false;
    state.templateInfo = null;
    
    document.getElementById('template-input').value = '';
    document.getElementById('template-upload-area').style.display = 'block';
    document.getElementById('template-info').style.display = 'none';
    document.getElementById('template-preview').style.display = 'none';
    document.getElementById('template-status').classList.remove('success');
    
    addLog('已移除PPT模板', 'info');
}

// ===== 图片上传处理 =====
async function handleImagesUpload(files) {
    if (!files || files.length === 0) return;
    
    // 支持的图片扩展名
    const allowedExts = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'];
    
    const formData = new FormData();
    for (let file of files) {
        // 通过扩展名判断是否为图片（某些系统可能无法正确识别 file.type）
        const ext = file.name.split('.').pop().toLowerCase();
        if (file.type.startsWith('image/') || allowedExts.includes(ext)) {
            formData.append('images', file);
        } else {
            addLog(`跳过非图片文件: ${file.name}`, 'warning');
        }
    }
    
    if (formData.getAll('images').length === 0) {
        addLog('没有有效的图片文件被选中', 'warning');
        return;
    }
    
    try {
        addLog(`正在上传 ${formData.getAll('images').length} 张图片...`, 'info');
        
        const response = await fetch('/api/upload/images', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            state.imagesUploaded = true;
            state.uploadedImages = data.images;
            
            // 更新UI
            document.getElementById('images-list').style.display = 'block';
            document.getElementById('images-count').textContent = data.count;
            document.getElementById('images-status').classList.add('success');
            
            // 显示图片列表
            const imagesUl = document.getElementById('images-ul');
            imagesUl.innerHTML = '';
            data.images.forEach(img => {
                const li = document.createElement('li');
                li.textContent = `${img.filename} (${img.width}x${img.height})`;
                imagesUl.appendChild(li);
            });
            
            addLog(`图片上传成功！共 ${data.count} 张`, 'success');
        } else {
            addLog(`上传失败：${data.error}`, 'error');
        }
    } catch (error) {
        addLog(`上传出错：${error.message}`, 'error');
    }
}

// ===== ZIP上传处理 =====
async function handleZipUpload(file) {
    if (!file) return;
    
    if (!file.name.endsWith('.zip')) {
        addLog('请上传 .zip 格式的压缩包', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('zip', file);
    
    try {
        addLog('正在上传并解压ZIP文件...', 'info');
        
        const response = await fetch('/api/upload/images', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            state.imagesUploaded = true;
            state.uploadedImages = data.images;
            
            // 更新UI
            document.getElementById('images-list').style.display = 'block';
            document.getElementById('images-count').textContent = data.count;
            document.getElementById('images-status').classList.add('success');
            
            // 显示图片列表
            const imagesUl = document.getElementById('images-ul');
            imagesUl.innerHTML = '';
            data.images.forEach(img => {
                const li = document.createElement('li');
                li.textContent = `${img.filename} (${img.width}x${img.height})`;
                imagesUl.appendChild(li);
            });
            
            addLog(`ZIP解压成功！共提取 ${data.count} 张图片`, 'success');
        } else {
            addLog(`上传失败：${data.error}`, 'error');
        }
    } catch (error) {
        addLog(`上传出错：${error.message}`, 'error');
    }
}

// ===== 价格表上传处理 =====
async function handlePriceUpload(file) {
    if (!file) return;
    
    const allowedExts = ['.csv', '.xlsx', '.xls'];
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!allowedExts.includes(ext)) {
        addLog('请上传 CSV 或 Excel 格式的价格表', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('price', file);
    
    try {
        addLog('正在上传价格表...', 'info');
        
        const response = await fetch('/api/upload/price', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            state.priceUploaded = true;
            state.priceInfo = data;
            
            // 更新UI
            document.getElementById('price-info').style.display = 'flex';
            document.getElementById('price-info').querySelector('.file-name').textContent = file.name;
            document.getElementById('price-upload-area').style.display = 'none';
            document.getElementById('price-status').classList.add('success');
            
            // 显示价格预览
            document.getElementById('price-preview').style.display = 'block';
            const priceList = document.getElementById('price-list');
            priceList.innerHTML = '';
            
            Object.entries(data.sample_data).forEach(([key, value]) => {
                const li = document.createElement('li');
                li.textContent = `${key}: ¥${value}`;
                priceList.appendChild(li);
            });
            
            addLog(`价格表上传成功！共 ${data.total_records} 条记录`, 'success');
        } else {
            addLog(`上传失败：${data.error}`, 'error');
        }
    } catch (error) {
        addLog(`上传出错：${error.message}`, 'error');
    }
}

// ===== 移除价格表 =====
function removePrice() {
    state.priceUploaded = false;
    state.priceInfo = null;
    
    document.getElementById('price-input').value = '';
    document.getElementById('price-upload-area').style.display = 'block';
    document.getElementById('price-info').style.display = 'none';
    document.getElementById('price-preview').style.display = 'none';
    document.getElementById('price-status').classList.remove('success');
    
    addLog('已移除价格表', 'info');
}

// ===== 生成产品手册 =====
async function generateManual() {
    // 验证必要文件
    if (!state.templateUploaded) {
        addLog('请先上传PPT模板', 'warning');
        document.getElementById('step1').scrollIntoView({ behavior: 'smooth' });
        return;
    }
    
    if (!state.imagesUploaded) {
        addLog('请先上传产品图片', 'warning');
        document.getElementById('step3').scrollIntoView({ behavior: 'smooth' });
        return;
    }
    
    // 获取导出格式
    const exportFormats = Array.from(document.querySelectorAll('input[name="export-format"]:checked'))
        .map(cb => cb.value);
    
    if (exportFormats.length === 0) {
        addLog('请至少选择一种导出格式', 'warning');
        return;
    }
    
    // 收集配置参数
    const config = {
        series_name: document.getElementById('series-name').value,
        series_font: document.getElementById('series-font').value,
        series_size: document.getElementById('series-size').value,
        series_bold: document.getElementById('series-bold').checked,
        series_italic: document.getElementById('series-italic').checked,
        
        product_info: document.getElementById('product-info').value,
        info_font: document.getElementById('info-font').value,
        info_size: document.getElementById('info-size').value,
        info_bold: document.getElementById('info-bold').checked,
        info_italic: document.getElementById('info-italic').checked,
        
        image_width: document.getElementById('image-width').value,
        image_height: document.getElementById('image-height').value,
        
        export_formats: exportFormats
    };
    
    // 显示加载
    document.getElementById('loading-overlay').style.display = 'flex';
    
    try {
        addLog('开始生成产品手册...', 'info');
        
        const response = await fetch('/api/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(config)
        });
        
        const data = await response.json();
        
        // 隐藏加载
        document.getElementById('loading-overlay').style.display = 'none';
        
        if (data.success) {
            addLog(`生成成功！共 ${data.product_count} 个产品`, 'success');
            
            // 显示下载链接
            const downloadCard = document.getElementById('download-card');
            const downloadList = document.getElementById('download-list');
            downloadList.innerHTML = '';
            
            const formatNames = {
                'pptx': { name: '可编辑PPT', icon: '📊' },
                'images_zip': { name: '分页图片ZIP', icon: '🖼️' },
                'long_image': { name: '合并长图', icon: '📜' }
            };
            
            Object.entries(data.downloads).forEach(([format, url]) => {
                const info = formatNames[format] || { name: format, icon: '📁' };
                
                const link = document.createElement('a');
                link.href = url;
                link.className = 'download-item';
                link.innerHTML = `<span class="icon">${info.icon}</span><span>${info.name}</span>`;
                link.download = '';
                
                downloadList.appendChild(link);
            });
            
            downloadCard.style.display = 'block';
            downloadCard.scrollIntoView({ behavior: 'smooth' });
        } else {
            addLog(`生成失败：${data.error}`, 'error');
        }
    } catch (error) {
        document.getElementById('loading-overlay').style.display = 'none';
        addLog(`生成出错：${error.message}`, 'error');
    }
}

// ===== 清除所有文件 =====
async function clearAll() {
    if (!confirm('确定要清除所有上传的文件吗？')) {
        return;
    }
    
    try {
        const response = await fetch('/api/clear', { method: 'POST' });
        const data = await response.json();
        
        if (data.success) {
            // 重置状态
            state.templateUploaded = false;
            state.imagesUploaded = false;
            state.priceUploaded = false;
            state.uploadedImages = [];
            state.templateInfo = null;
            state.priceInfo = null;
            
            // 重置UI
            removeTemplate();
            removePrice();
            
            document.getElementById('images-list').style.display = 'none';
            document.getElementById('images-ul').innerHTML = '';
            document.getElementById('images-status').classList.remove('success');
            document.getElementById('images-input').value = '';
            document.getElementById('zip-input').value = '';
            
            document.getElementById('download-card').style.display = 'none';
            document.getElementById('download-list').innerHTML = '';
            
            addLog('已清除所有文件', 'success');
        }
    } catch (error) {
        addLog(`清除失败：${error.message}`, 'error');
    }
}

// ===== 日志功能 =====
function addLog(message, type = 'info') {
    const container = document.getElementById('log-container');
    const time = new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    
    const logItem = document.createElement('div');
    logItem.className = `log-item ${type}`;
    logItem.innerHTML = `
        <span class="log-time">${time}</span>
        <span class="log-text">${message}</span>
    `;
    
    container.appendChild(logItem);
    container.scrollTop = container.scrollHeight;
}
