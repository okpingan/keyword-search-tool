const { ipcRenderer } = require('electron');

// 全局变量
let currentFile = null;
let excelData = null;
let searchResults = [];

// DOM元素
const elements = {
    selectFileBtn: document.getElementById('selectFileBtn'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileStatus: document.getElementById('fileStatus'),
    originalSheetSelect: document.getElementById('originalSheetSelect'),
    originalDataInfo: document.getElementById('originalDataInfo'),
    keywordInput: document.getElementById('keywordInput'),
    keywordPreview: document.getElementById('keywordPreview'),
    keywordCount: document.getElementById('keywordCount'),
    searchTarget: document.getElementById('searchTarget'),
    keywordCountSummary: document.getElementById('keywordCountSummary'),
    startSearchBtn: document.getElementById('startSearchBtn'),
    progressSection: document.getElementById('progressSection'),
    progressFill: document.getElementById('progressFill'),
    progressText: document.getElementById('progressText'),
    resultsSummary: document.getElementById('resultsSummary'),
    resultsTable: document.getElementById('resultsTable'),
    saveResultsBtn: document.getElementById('saveResultsBtn'),
    newSearchBtn: document.getElementById('newSearchBtn'),
    messageBox: document.getElementById('messageBox'),
    messageContent: document.getElementById('messageContent')
};

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

// 初始化事件监听器
function initializeEventListeners() {
    elements.selectFileBtn.addEventListener('click', selectFile);
    elements.originalSheetSelect.addEventListener('change', onOriginalSheetChange);
    elements.keywordInput.addEventListener('input', onKeywordInputChange);
    elements.startSearchBtn.addEventListener('click', startSearch);
    elements.saveResultsBtn.addEventListener('click', saveResults);
    elements.newSearchBtn.addEventListener('click', newSearch);
}

// 选择文件
async function selectFile() {
    try {
        showLoading('正在选择文件...');
        
        const filePath = await ipcRenderer.invoke('select-file');
        if (!filePath) {
            hideLoading();
            return;
        }

        currentFile = filePath;
        const fileName = filePath.split(/[\\\/]/).pop();
        
        // 显示文件信息
        elements.fileName.textContent = fileName;
        elements.fileStatus.textContent = '正在读取...';
        elements.fileStatus.className = 'status';
        elements.fileInfo.style.display = 'flex';
        
        // 读取Excel文件
        excelData = await ipcRenderer.invoke('read-excel', filePath);
        
        // 更新状态
        elements.fileStatus.textContent = '读取成功';
        elements.fileStatus.className = 'status success';
        
        // 填充工作表选择器
        populateSheetSelectors();
        
        // 显示步骤2
        document.getElementById('step2').style.display = 'block';
        
        hideLoading();
        showMessage('文件读取成功', 'success');
        
    } catch (error) {
        hideLoading();
        elements.fileStatus.textContent = '读取失败';
        elements.fileStatus.className = 'status error';
        showMessage(`文件读取失败: ${error.message}`, 'error');
    }
}

// 填充工作表选择器
function populateSheetSelectors() {
    if (!excelData || !excelData.sheets) return;
    
    // 清空现有选项
    elements.originalSheetSelect.innerHTML = '<option value="">请选择...</option>';
    
    // 添加工作表选项
    excelData.sheets.forEach(sheetName => {
        const option = new Option(sheetName, sheetName);
        elements.originalSheetSelect.appendChild(option);
    });
    
    // 自动选择可能的工作表
    autoSelectSheet();
}

// 自动选择工作表
function autoSelectSheet() {
    const sheets = excelData.sheets;
    
    // 尝试找到原始数据表（通常是第一个或包含"数据"、"原始"等关键词的表）
    let originalSheet = sheets[0];
    for (let sheet of sheets) {
        if (sheet.includes('数据') || sheet.includes('原始') || sheet.includes('data')) {
            originalSheet = sheet;
            break;
        }
    }
    
    elements.originalSheetSelect.value = originalSheet;
    
    // 触发change事件
    onOriginalSheetChange();
}

// 原始数据表选择变化
function onOriginalSheetChange() {
    const sheetName = elements.originalSheetSelect.value;
    if (!sheetName || !excelData.data[sheetName]) {
        elements.originalDataInfo.textContent = '';
        elements.searchTarget.textContent = '未选择';
        return;
    }
    
    const data = excelData.data[sheetName];
    const rowCount = Math.max(0, data.length - 1); // 减去标题行
    const colCount = data.length > 0 ? data[0].length : 0;
    
    elements.originalDataInfo.textContent = `${rowCount} 行数据，${colCount} 列`;
    elements.searchTarget.textContent = `${sheetName} (B列)`;
    
    // 显示步骤3
    document.getElementById('step3').style.display = 'block';
    
    checkCanProceed();
}

// 关键字输入变化
function onKeywordInputChange() {
    const inputText = elements.keywordInput.value.trim();
    const keywords = extractKeywordsFromInput(inputText);
    
    // 更新预览
    updateKeywordPreview(keywords);
    
    // 更新计数
    elements.keywordCount.textContent = `共 ${keywords.length} 个关键字`;
    elements.keywordCountSummary.textContent = keywords.length;
    
    checkCanProceed();
}

// 从输入框提取关键字
function extractKeywordsFromInput(inputText) {
    if (!inputText) return [];
    
    return inputText
        .split('\n')
        .map(k => k.trim())
        .filter(k => k && k.length > 0);
}

// 更新关键字预览
function updateKeywordPreview(keywords) {
    elements.keywordPreview.innerHTML = '';
    
    if (keywords.length === 0) {
        const span = document.createElement('span');
        span.className = 'no-keywords';
        span.textContent = '请输入关键字...';
        elements.keywordPreview.appendChild(span);
        return;
    }
    
    // 显示前10个关键字
    keywords.slice(0, 10).forEach((keyword, index) => {
        const span = document.createElement('span');
        span.className = 'keyword-tag';
        span.textContent = keyword;
        elements.keywordPreview.appendChild(span);
    });
    
    // 如果超过10个，显示省略提示
    if (keywords.length > 10) {
        const span = document.createElement('span');
        span.className = 'keyword-more';
        span.textContent = `...还有 ${keywords.length - 10} 个`;
        elements.keywordPreview.appendChild(span);
    }
}

// 检查是否可以进行下一步
function checkCanProceed() {
    const originalSheet = elements.originalSheetSelect.value;
    const inputText = elements.keywordInput.value.trim();
    const keywords = extractKeywordsFromInput(inputText);
    
    if (originalSheet && keywords.length > 0) {
        document.getElementById('step4').style.display = 'block';
    } else {
        document.getElementById('step4').style.display = 'none';
    }
}

// 开始检索
async function startSearch() {
    try {
        const originalSheet = elements.originalSheetSelect.value;
        const inputText = elements.keywordInput.value.trim();
        
        if (!originalSheet) {
            showMessage('请先选择原始数据表', 'error');
            return;
        }
        
        if (!inputText) {
            showMessage('请输入关键字', 'error');
            return;
        }
        
        const keywords = extractKeywordsFromInput(inputText);
        
        if (keywords.length === 0) {
            showMessage('请输入至少一个关键字', 'error');
            return;
        }
        
        const originalData = excelData.data[originalSheet];
        
        // 显示进度条
        showProgress();
        updateProgress(10, '保存关键字到Excel...');
        
        // 禁用按钮
        elements.startSearchBtn.disabled = true;
        
        // 保存关键字到Excel文件
        try {
            const saveKeywordResult = await ipcRenderer.invoke('save-keywords-to-excel', {
                filePath: currentFile,
                keywords: keywords
            });
            
            if (saveKeywordResult.success) {
                console.log(saveKeywordResult.message);
                updateProgress(30, '关键字已保存，开始检索...');
            } else {
                console.warn('关键字保存失败:', saveKeywordResult.message);
                updateProgress(30, '开始检索...');
            }
        } catch (keywordSaveError) {
            console.warn('保存关键字时出错:', keywordSaveError);
            updateProgress(30, '开始检索...');
        }
        
        // 执行检索
        updateProgress(50, '正在检索匹配项...');
        
        searchResults = await ipcRenderer.invoke('search-keywords', {
            originalData: originalData,
            keywords: keywords,
            sheetName: originalSheet
        });
        
        updateProgress(100, '检索完成');
        
        // 显示结果
        displayResults();
        
        // 显示步骤5
        document.getElementById('step5').style.display = 'block';
        
        showMessage(`检索完成，找到 ${searchResults.length} 条匹配记录`, 'success');
        
    } catch (error) {
        showMessage(`检索失败: ${error.message}`, 'error');
    } finally {
        elements.startSearchBtn.disabled = false;
        hideProgress();
    }
}

// 显示进度条
function showProgress() {
    elements.progressSection.style.display = 'block';
    elements.progressFill.style.width = '0%';
}

// 更新进度
function updateProgress(percent, text) {
    elements.progressFill.style.width = `${percent}%`;
    elements.progressText.textContent = text;
}

// 隐藏进度条
function hideProgress() {
    setTimeout(() => {
        elements.progressSection.style.display = 'none';
    }, 1000);
}

// 显示检索结果
function displayResults() {
    if (!searchResults || searchResults.length === 0) {
        elements.resultsSummary.innerHTML = '<p>未找到匹配的结果</p>';
        elements.resultsTable.innerHTML = '';
        return;
    }
    
    // 显示统计信息
    displayResultsSummary();
    
    // 显示结果表格
    displayResultsTable();
}

// 显示结果统计
function displayResultsSummary() {
    const totalResults = searchResults.length;
    const uniqueKeywords = [...new Set(
        searchResults.flatMap(r => r.关键字.split(',').map(k => k.trim()))
    )].length;
    
    elements.resultsSummary.innerHTML = `
        <div class="stat-item">
            <span class="stat-value">${totalResults}</span>
            <div class="stat-label">匹配记录</div>
        </div>
        <div class="stat-item">
            <span class="stat-value">${uniqueKeywords}</span>
            <div class="stat-label">匹配关键字</div>
        </div>
    `;
}

// 显示结果表格
function displayResultsTable() {
    const previewResults = searchResults.slice(0, 10);
    
    if (previewResults.length === 0) {
        elements.resultsTable.innerHTML = '<p>暂无数据</p>';
        return;
    }
    
    // 获取原始数据的列名
    const originalHeaders = Object.keys(previewResults[0].原始数据);
    
    let tableHTML = `
        <table>
            <thead>
                <tr>
                    <th>行号</th>
                    <th>关键字</th>
                    ${originalHeaders.map(header => `<th>${header}</th>`).join('')}
                </tr>
            </thead>
            <tbody>
    `;
    
    previewResults.forEach(result => {
        tableHTML += `
            <tr>
                <td>${result.序号}</td>
                <td><span class="keyword-highlight">${result.关键字}</span></td>
                ${originalHeaders.map(header => {
                    let cellValue = result.原始数据[header] || '';
                    return `<td title="${cellValue}">${cellValue}</td>`;
                }).join('')}
            </tr>
        `;
    });
    
    tableHTML += `
            </tbody>
        </table>
    `;
    
    if (searchResults.length > 10) {
        tableHTML += `<p class="preview-note">显示前10条结果，共${searchResults.length}条记录</p>`;
    }
    
    elements.resultsTable.innerHTML = tableHTML;
}

// 保存结果
async function saveResults() {
    if (!searchResults || searchResults.length === 0) {
        showMessage('没有结果可以保存', 'error');
        return;
    }
    
    try {
        showLoading('正在保存结果...');
        
        // 获取原始数据的列名
        const originalHeaders = Object.keys(searchResults[0].原始数据);
        
        const savedPath = await ipcRenderer.invoke('save-results', {
            results: searchResults,
            originalHeaders: originalHeaders
        });
        
        if (savedPath) {
            showMessage(`结果已保存到: ${savedPath}`, 'success');
        }
        
    } catch (error) {
        showMessage(`保存失败: ${error.message}`, 'error');
    } finally {
        hideLoading();
    }
}

// 重新检索
function newSearch() {
    // 清空结果
    searchResults = [];
    
    // 隐藏步骤4和步骤5
    document.getElementById('step4').style.display = 'none';
    document.getElementById('step5').style.display = 'none';
    
    // 清空输入框
    elements.keywordInput.value = '';
    
    // 重置预览
    updateKeywordPreview([]);
    elements.keywordCount.textContent = '';
    elements.keywordCountSummary.textContent = '0';
    
    // 清空结果显示
    elements.resultsSummary.innerHTML = '';
    elements.resultsTable.innerHTML = '';
    
    showMessage('已重置，请重新输入关键字', 'info');
}

// 显示消息
function showMessage(message, type = 'info') {
    elements.messageContent.textContent = message;
    elements.messageBox.className = `message-box ${type}`;
    elements.messageBox.style.display = 'block';
    
    // 添加显示动画
    setTimeout(() => {
        elements.messageBox.classList.add('show');
    }, 10);
    
    // 3秒后自动隐藏
    setTimeout(() => {
        hideMessage();
    }, 3000);
}

// 隐藏消息
function hideMessage() {
    elements.messageBox.classList.remove('show');
    setTimeout(() => {
        elements.messageBox.style.display = 'none';
    }, 300);
}

// 显示加载状态
function showLoading(message) {
    showMessage(message, 'info');
}

// 隐藏加载状态
function hideLoading() {
    hideMessage();
}