const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    icon: path.join(__dirname, 'icon.ico'),
    title: '关键字检索工具',
    show: false
  });

  mainWindow.loadFile('index.html');
  
  // 窗口准备好后显示
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // 开发模式下打开调试工具
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// 处理文件选择
ipcMain.handle('select-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '选择Excel文件',
    filters: [
      { name: 'Excel文件', extensions: ['xlsx', 'xls'] }
    ],
    properties: ['openFile']
  });
  
  if (!result.canceled && result.filePaths.length > 0) {
    return result.filePaths[0];
  }
  return null;
});

// 处理Excel文件读取
ipcMain.handle('read-excel', async (event, filePath) => {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    
    let result = {
      sheets: sheetNames,
      data: {}
    };
    
    // 读取所有工作表
    sheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      result.data[sheetName] = jsonData;
    });
    
    return result;
  } catch (error) {
    throw new Error(`读取Excel文件失败: ${error.message}`);
  }
});

// 处理关键字检索
ipcMain.handle('search-keywords', async (event, { originalData, keywords, sheetName }) => {
  try {
    const results = [];
    
    // 获取原始数据的标题行
    const headers = originalData[0] || [];
    const dataRows = originalData.slice(1);
    
    // 找到B列（数据集列）的索引
    const datasetColumnIndex = 1; // B列索引为1
    
    // 准备所有关键字的列表（扁平化处理）
    const allKeywords = [];
    keywords.forEach(keyword => {
      if (typeof keyword === 'string') {
        // 如果是字符串，按逗号分割
        const keywordList = keyword.split(',').map(k => k.trim()).filter(k => k);
        allKeywords.push(...keywordList);
      } else {
        // 如果已经是数组形式，直接添加
        allKeywords.push(keyword);
      }
    });
    
    // 遍历每一行数据
    dataRows.forEach((row, rowIndex) => {
      const cellValue = row[datasetColumnIndex] || '';
      const cellText = cellValue.toString().toLowerCase();
      
      // 收集该行匹配的所有关键字
      const matchedKeywords = [];
      allKeywords.forEach(keyword => {
        if (cellText.includes(keyword.toLowerCase())) {
          // 避免重复添加相同的关键字
          if (!matchedKeywords.includes(keyword)) {
            matchedKeywords.push(keyword);
          }
        }
      });
      
      // 如果有匹配的关键字，创建一条结果记录
      if (matchedKeywords.length > 0) {
        const resultRow = {
          序号: rowIndex + 2, // Excel行号（从第2行开始）
          关键字: matchedKeywords.join(','), // 合并所有匹配的关键字
          匹配字段: '数据集',
          原始数据: {}
        };
        
        // 添加原始数据的所有列
        headers.forEach((header, index) => {
          if (header) {
            resultRow.原始数据[header] = row[index] || '';
          }
        });
        
        results.push(resultRow);
      }
    });
    
    return results;
  } catch (error) {
    throw new Error(`检索失败: ${error.message}`);
  }
});

// 处理保存结果
ipcMain.handle('save-results', async (event, { results, originalHeaders }) => {
  try {
    const saveResult = await dialog.showSaveDialog(mainWindow, {
      title: '保存检索结果',
      defaultPath: `检索结果_${new Date().toISOString().slice(0, 10)}.xlsx`,
      filters: [
        { name: 'Excel文件', extensions: ['xlsx'] }
      ]
    });
    
    if (saveResult.canceled) {
      return null;
    }
    
    // 准备结果数据
    const resultData = [];
    
    // 标题行 - 直接使用原始标题，不添加额外的列
    resultData.push(originalHeaders);
    
    // 数据行 - 直接使用原始数据，保持原有的行号
    results.forEach(result => {
      const row = [];
      
      // 添加原始数据的所有列
      originalHeaders.forEach(header => {
        row.push(result.原始数据[header] || '');
      });
      
      resultData.push(row);
    });
    
    // 创建工作簿
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(resultData);
    
    // 设置列宽
    const colWidths = originalHeaders.map(header => ({ wch: Math.max(header.length, 15) }));
    worksheet['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(workbook, worksheet, '检索结果');
    
    // 保存文件
    XLSX.writeFile(workbook, saveResult.filePath);
    
    return saveResult.filePath;
  } catch (error) {
    throw new Error(`保存失败: ${error.message}`);
  }
});

// 新增：将关键字写入Excel文件的"关键字"工作表
ipcMain.handle('save-keywords-to-excel', async (event, { filePath, keywords }) => {
  try {
    const workbook = XLSX.readFile(filePath);
    
    // 查找是否已有"关键字"表
    let targetSheetName = workbook.SheetNames.find(name => 
      ['关键字', 'keywords', '关键词', '搜索'].some(key => 
        name.toLowerCase().includes(key.toLowerCase())
      )
    );
    
    // 如果没有，则新建
    if (!targetSheetName) {
      targetSheetName = '关键字';
    }
    
    // 构造新数据：标题 + 一行一个关键字
    const newData = [['关键字'], ...keywords.map(k => [k])];
    
    // 创建新工作表
    const newSheet = XLSX.utils.aoa_to_sheet(newData);
    
    // 设置列宽
    newSheet['!cols'] = [{ wch: 20 }];
    
    // 添加或替换工作表
    workbook.Sheets[targetSheetName] = newSheet;
    
    // 如果是新工作表，添加到工作表名称列表
    if (!workbook.SheetNames.includes(targetSheetName)) {
      workbook.SheetNames.push(targetSheetName);
    }
    
    // 保存文件
    XLSX.writeFile(workbook, filePath);
    
    return { 
      success: true, 
      message: `已将 ${keywords.length} 个关键字保存到 "${targetSheetName}" 工作表` 
    };
  } catch (error) {
    return { 
      success: false, 
      message: `保存关键字失败: ${error.message}` 
    };
  }
});