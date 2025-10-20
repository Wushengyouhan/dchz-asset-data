const DatabaseManager = require('./database');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const config = require('./config');

class ContractDataProcessor {
  constructor() {
    this.dbManager = new DatabaseManager(config, 'redSystem');
  }

  /**
   * 根据原资产编码查询新资产信息
   * @param {string} originalAssetCode - 原资产编码
   * @returns {Object|null} 新资产信息或null
   */
  async getNewAssetInfo(originalAssetCode) {
    try {
      const query = `
        SELECT 
          a2.AS_CODE AS 新资产编码,
          a2.AS_NAME AS 新资产名称,
          a2.AS_LV AS 新资产等级,
          a2.OPERATING AS 新资产类型,
          a2.AS_USABLE_AREA AS 新资产可用面积 
        FROM as_asset a1
        INNER JOIN as_asset a2 ON a1.NEW_AS_CODE = a2.AS_CODE
        WHERE a1.AS_CODE = ?
          AND a1.U_DELETE = 1
          AND a2.U_DELETE = 1
      `;

      const results = await this.dbManager.query(query, [originalAssetCode]);
      
      if (results.length > 0) {
        console.log(`✅ 查询成功: ${originalAssetCode} -> ${results[0].新资产编码}`);
        return results[0];
      } else {
        console.log(`❌ 查询失败: ${originalAssetCode} 未找到新资产信息`);
        return null;
      }
    } catch (error) {
      console.error(`❌ 查询新资产信息失败 (原资产编码: ${originalAssetCode}):`, error.message);
      return null;
    }
  }

  /**
   * 读取十堰西资产编码Excel文件
   * @param {string} filePath - Excel文件路径
   * @returns {Array} 十堰西资产数据列表
   */
  readShiyanWestAssets(filePath) {
    try {
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0]; // 使用第一个工作表
      const worksheet = workbook.Sheets[sheetName];
      
      console.log(`📋 使用工作表: ${sheetName}`);
      
      // 将工作表转换为JSON格式
      const allData = XLSX.utils.sheet_to_json(worksheet);
      
      console.log(`✅ 成功读取Excel文件: ${filePath}`);
      console.log(`📊 共读取 ${allData.length} 行数据`);
      
      // 显示所有可用的列名
      if (allData.length > 0) {
        console.log('📋 可用的列名:', Object.keys(allData[0]));
      }
      
      // 由于这个文件专门是十堰西资产编码，不需要筛选，直接返回所有数据
      console.log(`🔍 十堰西资产编码数据: ${allData.length} 条`);
      
      return allData;
    } catch (error) {
      console.error('❌ 读取Excel文件失败:', error.message);
      throw error;
    }
  }

  /**
   * 处理单个资产编码，查询新资产信息
   * @param {string} originalAssetCode - 原资产编码
   * @returns {Object} 处理结果
   */
  async processAssetCode(originalAssetCode) {
    console.log(`🔍 正在处理十堰西资产经营中心原资产编码: ${originalAssetCode}`);
    
    const newAssetInfo = await this.getNewAssetInfo(originalAssetCode);
    
    if (newAssetInfo) {
      console.log(`✅ 找到新资产信息: ${newAssetInfo.新资产编码} - ${newAssetInfo.新资产名称}`);
      return {
        原资产编码: originalAssetCode,
        新资产编码: newAssetInfo.新资产编码,
        新资产名称: newAssetInfo.新资产名称,
        新资产等级: newAssetInfo.新资产等级,
        新资产类型: newAssetInfo.新资产类型,
        新资产可用面积: newAssetInfo.新资产可用面积,
        状态: '已找到'
      };
    } else {
      console.log(`⚠️  未找到新资产信息: ${originalAssetCode}`);
      return {
        原资产编码: originalAssetCode,
        新资产编码: '000',
        新资产名称: '',
        新资产等级: '',
        新资产类型: '',
        新资产可用面积: '',
        状态: '未找到'
      };
    }
  }

  /**
   * 批量处理十堰西资产数据
   * @param {Array} assetData - 十堰西资产数据列表
   * @returns {Array} 处理结果列表（包含原数据和新资产信息）
   */
  async processShiyanWestAssets(assetData) {
    const results = [];
    
    console.log(`🚀 开始批量处理 ${assetData.length} 个十堰西资产...`);
    
    for (let i = 0; i < assetData.length; i++) {
      const row = assetData[i];
      console.log(`\n📋 处理进度: ${i + 1}/${assetData.length}`);
      
      try {
        // 获取原资产编码（需要根据实际Excel结构调整）
        const originalAssetCode = row['原资产编码'] || row['资产编码'] || Object.values(row)[0];
        
        if (!originalAssetCode) {
          console.log(`⚠️  跳过无资产编码的行: ${JSON.stringify(row)}`);
          results.push({
            ...row,
            新资产编码: '000',
            新资产名称: '',
            新资产等级: '',
            新资产类型: '',
            新资产可用面积: '',
            状态: '无资产编码'
          });
          continue;
        }
        
        console.log(`🔍 正在处理十堰西资产经营中心原资产编码: ${originalAssetCode}`);
        
        const newAssetInfo = await this.getNewAssetInfo(originalAssetCode);
        
        if (newAssetInfo) {
          console.log(`✅ 找到新资产信息: ${newAssetInfo.新资产编码} - ${newAssetInfo.新资产名称}`);
          results.push({
            ...row,
            新资产编码: newAssetInfo.新资产编码,
            新资产名称: newAssetInfo.新资产名称,
            新资产等级: newAssetInfo.新资产等级,
            新资产类型: newAssetInfo.新资产类型,
            新资产可用面积: newAssetInfo.新资产可用面积,
            状态: '已找到'
          });
        } else {
          console.log(`⚠️  未找到新资产信息: ${originalAssetCode}`);
          results.push({
            ...row,
            新资产编码: '000',
            新资产名称: '',
            新资产等级: '',
            新资产类型: '',
            新资产可用面积: '',
            状态: '未找到'
          });
        }
        
        // 添加延迟避免数据库压力过大
        if (i < assetData.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      } catch (error) {
        console.error(`❌ 处理资产失败: ${JSON.stringify(row)}`, error.message);
        results.push({
          ...row,
          新资产编码: '000',
          新资产名称: '',
          新资产等级: '',
          新资产类型: '',
          新资产可用面积: '',
          状态: '处理失败'
        });
      }
    }
    
    return results;
  }

  /**
   * 生成处理结果Excel文件（直接复制原Excel，在每行后添加5列新资产信息）
   * @param {Array} results - 处理结果
   * @param {string} outputPath - 输出文件路径
   * @param {string} originalFilePath - 原Excel文件路径
   */
  generateResultExcel(results, outputPath, originalFilePath) {
    try {
      if (results.length === 0) {
        console.log('⚠️  没有数据需要生成Excel文件');
        return;
      }
      
      // 读取原Excel文件，使用第三个tab
      const originalWorkbook = XLSX.readFile(originalFilePath);
      
      // 使用第一个工作表（十堰西资产编码文件）
      const targetSheetName = originalWorkbook.SheetNames[0];
      console.log(`📋 使用原Excel的工作表格式: ${targetSheetName}`);
      const originalWorksheet = originalWorkbook.Sheets[targetSheetName];
      
      // 获取原Excel的范围
      const range = XLSX.utils.decode_range(originalWorksheet['!ref']);
      
      // 创建新工作簿
      const newWorkbook = XLSX.utils.book_new();
      
      // 复制原工作表
      const newWorksheet = XLSX.utils.aoa_to_sheet([]);
      
      // 复制原Excel的所有数据
      for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          if (originalWorksheet[cellAddress]) {
            newWorksheet[cellAddress] = { ...originalWorksheet[cellAddress] };
          }
        }
      }
      
      // 添加新的5列表头
      const newHeaderCol = range.e.c + 1;
      const newHeaders = ['新资产编码', '新资产名称', '新资产等级', '新资产类型', '新资产可用面积'];
      
      newHeaders.forEach((header, index) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: newHeaderCol + index });
        newWorksheet[cellAddress] = { v: header, t: 's' };
      });
      
      // 为每行添加新资产信息
      results.forEach((result, index) => {
        const row = index + 1; // 从第2行开始（第1行是表头）
        const newDataCol = range.e.c + 1;
        
        // 添加5列新资产信息
        const newData = [
          result.新资产编码 || '',
          result.新资产名称 || '',
          result.新资产等级 || '',
          result.新资产类型 || '',
          result.新资产可用面积 || ''
        ];
        
        newData.forEach((value, colIndex) => {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: newDataCol + colIndex });
          newWorksheet[cellAddress] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
        });
      });
      
      // 更新工作表范围
      newWorksheet['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: range.e.r, c: range.e.c + 5 }
      });
      
      // 设置列宽
      const columnWidths = [];
      for (let col = 0; col <= range.e.c + 5; col++) {
        if (col > range.e.c) {
          // 新添加的列
          columnWidths.push({ wch: 15 });
        } else {
          // 原列保持默认宽度
          columnWidths.push({ wch: 15 });
        }
      }
      newWorksheet['!cols'] = columnWidths;
      
      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, targetSheetName);
      
      // 写入文件
      XLSX.writeFile(newWorkbook, outputPath);
      
      console.log(`✅ 结果Excel文件生成成功: ${outputPath}`);
      console.log(`📊 共处理 ${results.length} 条记录`);
      
      // 统计信息
      const foundCount = results.filter(r => r.状态 === '已找到').length;
      const notFoundCount = results.filter(r => r.状态 === '未找到').length;
      const failedCount = results.filter(r => r.状态 === '处理失败').length;
      const noCodeCount = results.filter(r => r.状态 === '无资产编码').length;
      
      console.log(`\n📈 处理统计:`);
      console.log(`   已找到: ${foundCount} 条`);
      console.log(`   未找到: ${notFoundCount} 条`);
      console.log(`   处理失败: ${failedCount} 条`);
      console.log(`   无资产编码: ${noCodeCount} 条`);
      
    } catch (error) {
      console.error('❌ 生成结果Excel文件失败:', error.message);
      throw error;
    }
  }

  /**
   * 处理Excel文件中的十堰西资产数据
   * @param {string} inputFilePath - 输入Excel文件路径
   * @param {string} outputFilePath - 输出Excel文件路径
   */
  async processExcelFile(inputFilePath, outputFilePath) {
    try {
      console.log('🚀 开始处理十堰西资产经营中心数据...');
      
      // 1. 连接数据库
      await this.dbManager.connect();
      
      // 2. 读取Excel文件并筛选十堰西数据
      console.log('📖 正在读取Excel文件并筛选十堰西资产...');
      const shiyanWestData = this.readShiyanWestAssets(inputFilePath);
      
      if (shiyanWestData.length === 0) {
        console.log('❌ 未找到十堰西资产经营中心的数据，程序结束');
        return;
      }
      
      console.log(`📋 找到 ${shiyanWestData.length} 条十堰西资产数据`);
      
      // 3. 批量处理十堰西资产数据
      const results = await this.processShiyanWestAssets(shiyanWestData);
      
      // 4. 生成结果Excel文件（直接复制原Excel并添加新列）
      console.log('📊 正在生成结果Excel文件...');
      this.generateResultExcel(results, outputFilePath, inputFilePath);
      
      console.log('✅ 处理完成！');
      
    } catch (error) {
      console.error('❌ 处理过程中发生错误:', error.message);
      throw error;
    } finally {
      // 关闭数据库连接
      await this.dbManager.close();
    }
  }
}

module.exports = ContractDataProcessor;
