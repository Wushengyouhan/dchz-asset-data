const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class AssetComparisonGenerator {
  constructor(config) {
    this.config = config;
  }

  /**
   * 读取Excel文件并解析为JSON数据
   * @param {string} filePath - Excel文件路径
   * @returns {Array} 解析后的数据数组
   */
  readExcelFile(filePath) {
    try {
      if (!fs.existsSync(filePath)) {
        throw new Error(`文件不存在: ${filePath}`);
      }

      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0]; // 读取第一个工作表
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet);

      console.log(`✅ 成功读取文件: ${filePath}，共 ${data.length} 条记录`);
      return data;
    } catch (error) {
      console.error(`❌ 读取文件失败: ${filePath}`, error.message);
      throw error;
    }
  }

  /**
   * 查找红色系统中匹配的资产
   * @param {Object} blueAsset - 蓝色系统资产
   * @param {Array} redAssets - 红色系统资产列表
   * @returns {Object|null} 匹配的红色系统资产
   */
  findMatchingRedAsset(blueAsset, redAssets) {
    // 第一步：通过OLD_AS_CODE匹配
    const matchByOldCode = redAssets.find(redAsset => 
      redAsset['OLD_AS_CODE'] && redAsset['OLD_AS_CODE'] === blueAsset['资产编码']
    );

    if (matchByOldCode) {
      console.log(`   ✅ 通过OLD_AS_CODE匹配: ${blueAsset['资产编码']} -> ${matchByOldCode['资产编码']}`);
      return matchByOldCode;
    }

    // 第二步：通过资产名称匹配
    const matchByName = redAssets.find(redAsset => 
      redAsset['资产名称'] && redAsset['资产名称'] === blueAsset['资产名称']
    );

    if (matchByName) {
      console.log(`   ✅ 通过资产名称匹配: ${blueAsset['资产名称']} -> ${matchByName['资产编码']}`);
      return matchByName;
    }

    console.log(`   ⚠️  未找到匹配: ${blueAsset['资产编码']} - ${blueAsset['资产名称']}`);
    return null;
  }

  /**
   * 查找蓝色系统中匹配的资产
   * @param {Object} redAsset - 红色系统资产
   * @param {Array} blueAssets - 蓝色系统资产列表
   * @returns {Object|null} 匹配的蓝色系统资产
   */
  findMatchingBlueAsset(redAsset, blueAssets) {
    // 第一步：通过OLD_AS_CODE匹配
    const matchByOldCode = blueAssets.find(blueAsset => 
      redAsset['OLD_AS_CODE'] && redAsset['OLD_AS_CODE'] === blueAsset['资产编码']
    );

    if (matchByOldCode) {
      console.log(`   ✅ 通过OLD_AS_CODE匹配: ${redAsset['资产编码']} -> ${matchByOldCode['资产编码']}`);
      return matchByOldCode;
    }

    // 第二步：通过资产名称匹配
    const matchByName = blueAssets.find(blueAsset => 
      redAsset['资产名称'] && redAsset['资产名称'] === blueAsset['资产名称']
    );

    if (matchByName) {
      console.log(`   ✅ 通过资产名称匹配: ${redAsset['资产名称']} -> ${matchByName['资产编码']}`);
      return matchByName;
    }

    console.log(`   ⚠️  未找到匹配: ${redAsset['资产编码']} - ${redAsset['资产名称']}`);
    return null;
  }

  /**
   * 以蓝色系统为主线生成对照表数据
   * @param {Array} blueAssets - 蓝色系统资产数据
   * @param {Array} redAssets - 红色系统资产数据
   * @returns {Array} 对照表数据
   */
  generateBlueSystemComparisonData(blueAssets, redAssets) {
    console.log('🔍 开始生成蓝色系统对照表...');
    console.log(`   蓝色系统资产: ${blueAssets.length} 条`);
    console.log(`   红色系统资产: ${redAssets.length} 条`);

    const comparisonData = [];
    let matchedCount = 0;
    let blueUnmatchedCount = 0;
    const usedRedAssets = new Set(); // 记录已使用的红色系统资产

    // 按照蓝色系统的顺序处理
    for (let i = 0; i < blueAssets.length; i++) {
      const blueAsset = blueAssets[i];
      console.log(`   ${i + 1}/${blueAssets.length} 处理蓝色资产: ${blueAsset['资产编码']} - ${blueAsset['资产名称']}`);

      const matchingRedAsset = this.findMatchingRedAsset(blueAsset, redAssets);

      if (matchingRedAsset) {
        matchedCount++;
        usedRedAssets.add(matchingRedAsset['资产编码']);
        
        // 创建对照行
        const comparisonRow = {
          // 蓝色系统字段
          '蓝色系统_资产编码': blueAsset['资产编码'],
          '蓝色系统_资产名称': blueAsset['资产名称'],
          '蓝色系统_资产等级': blueAsset['资产等级'],
          '蓝色系统_建筑面积': blueAsset['建筑面积'],
          '蓝色系统_租赁面积': blueAsset['租赁面积'],
          '蓝色系统_资产分类': blueAsset['资产分类'] || '',

          // 红色系统字段
          '红色系统_资产编码': matchingRedAsset['资产编码'],
          '红色系统_资产名称': matchingRedAsset['资产名称'],
          '红色系统_资产等级': matchingRedAsset['资产等级'],
          '红色系统_建筑面积': matchingRedAsset['建筑面积'],
          '红色系统_租赁面积': matchingRedAsset['租赁面积'],
          '红色系统_资产分类': matchingRedAsset['资产分类'] || '',

          // 匹配信息
          '匹配状态': '已匹配',
          '匹配方式': matchingRedAsset['OLD_AS_CODE'] === blueAsset['资产编码'] ? '编码' : '名称'
        };

        comparisonData.push(comparisonRow);
      } else {
        blueUnmatchedCount++;
        // 创建蓝色系统未匹配行（只包含蓝色系统数据）
        const unmatchedRow = {
          // 蓝色系统字段
          '蓝色系统_资产编码': blueAsset['资产编码'],
          '蓝色系统_资产名称': blueAsset['资产名称'],
          '蓝色系统_资产等级': blueAsset['资产等级'],
          '蓝色系统_建筑面积': blueAsset['建筑面积'],
          '蓝色系统_租赁面积': blueAsset['租赁面积'],
          '蓝色系统_资产分类': blueAsset['资产分类'] || '',

          // 红色系统字段（空值）
          '红色系统_资产编码': '',
          '红色系统_资产名称': '',
          '红色系统_资产等级': '',
          '红色系统_建筑面积': '',
          '红色系统_租赁面积': '',
          '红色系统_资产分类': '',

          // 匹配信息
          '匹配状态': '未匹配',
          '匹配方式': ''
        };

        comparisonData.push(unmatchedRow);
      }
    }

    console.log(`\n📊 蓝色系统对照表统计:`);
    console.log(`   已匹配: ${matchedCount} 条`);
    console.log(`   蓝色系统未匹配: ${blueUnmatchedCount} 条`);
    console.log(`   总计: ${comparisonData.length} 条`);

    return comparisonData;
  }

  /**
   * 以红色系统为主线生成对照表数据（红色系统字段在前）
   * @param {Array} blueAssets - 蓝色系统资产数据
   * @param {Array} redAssets - 红色系统资产数据
   * @returns {Array} 对照表数据
   */
  generateRedSystemComparisonData(blueAssets, redAssets) {
    console.log('🔍 开始生成红色系统对照表...');
    console.log(`   蓝色系统资产: ${blueAssets.length} 条`);
    console.log(`   红色系统资产: ${redAssets.length} 条`);

    const comparisonData = [];
    let matchedCount = 0;
    let redUnmatchedCount = 0;
    const usedBlueAssets = new Set(); // 记录已使用的蓝色系统资产

    // 按照红色系统的顺序处理，保持层级结构
    for (let i = 0; i < redAssets.length; i++) {
      const redAsset = redAssets[i];
      console.log(`   ${i + 1}/${redAssets.length} 处理红色资产: ${redAsset['资产编码']} - ${redAsset['资产名称']}`);

      // 查找匹配的蓝色系统资产
      const matchingBlueAsset = this.findMatchingRedAsset(redAsset, blueAssets);

      if (matchingBlueAsset) {
        matchedCount++;
        usedBlueAssets.add(matchingBlueAsset['资产编码']);
        
        // 创建对照行（红色系统在前，蓝色系统在后）
        const comparisonRow = {
          // 红色系统字段
          '红色系统_资产编码': redAsset['资产编码'],
          '红色系统_资产名称': redAsset['资产名称'],
          '红色系统_资产等级': redAsset['资产等级'],
          '红色系统_建筑面积': redAsset['建筑面积'],
          '红色系统_租赁面积': redAsset['租赁面积'],
          '红色系统_资产分类': redAsset['资产分类'] || '',

          // 蓝色系统字段
          '蓝色系统_资产编码': matchingBlueAsset['资产编码'],
          '蓝色系统_资产名称': matchingBlueAsset['资产名称'],
          '蓝色系统_资产等级': matchingBlueAsset['资产等级'],
          '蓝色系统_建筑面积': matchingBlueAsset['建筑面积'],
          '蓝色系统_租赁面积': matchingBlueAsset['租赁面积'],
          '蓝色系统_资产分类': matchingBlueAsset['资产分类'] || '',

          // 匹配信息
          '匹配状态': '已匹配',
          '匹配方式': redAsset['OLD_AS_CODE'] === matchingBlueAsset['资产编码'] ? '编码' : '名称'
        };

        comparisonData.push(comparisonRow);
      } else {
        redUnmatchedCount++;
        // 创建红色系统未匹配行（红色系统在前，蓝色系统在后）
        const redUnmatchedRow = {
          // 红色系统字段
          '红色系统_资产编码': redAsset['资产编码'],
          '红色系统_资产名称': redAsset['资产名称'],
          '红色系统_资产等级': redAsset['资产等级'],
          '红色系统_建筑面积': redAsset['建筑面积'],
          '红色系统_租赁面积': redAsset['租赁面积'],
          '红色系统_资产分类': redAsset['资产分类'] || '',

          // 蓝色系统字段（空值）
          '蓝色系统_资产编码': '',
          '蓝色系统_资产名称': '',
          '蓝色系统_资产等级': '',
          '蓝色系统_建筑面积': '',
          '蓝色系统_租赁面积': '',
          '蓝色系统_资产分类': '',

          // 匹配信息
          '匹配状态': '未匹配',
          '匹配方式': ''
        };

        comparisonData.push(redUnmatchedRow);
      }
    }

    console.log(`\n📊 红色系统对照表统计:`);
    console.log(`   已匹配: ${matchedCount} 条`);
    console.log(`   红色系统未匹配: ${redUnmatchedCount} 条`);
    console.log(`   总计: ${comparisonData.length} 条`);

    return comparisonData;
  }

  /**
   * 生成资产对照表数据（兼容旧版本）
   * @param {Array} blueAssets - 蓝色系统资产数据
   * @param {Array} redAssets - 红色系统资产数据
   * @returns {Array} 对照表数据
   */
  generateComparisonData(blueAssets, redAssets) {
    console.log('🔍 开始生成资产对照表（按红色系统层级结构）...');
    console.log(`   蓝色系统资产: ${blueAssets.length} 条`);
    console.log(`   红色系统资产: ${redAssets.length} 条`);

    const comparisonData = [];
    let matchedCount = 0;
    let blueUnmatchedCount = 0;
    let redUnmatchedCount = 0;
    const usedBlueAssets = new Set(); // 记录已使用的蓝色系统资产

    // 按照红色系统的顺序处理，保持层级结构
    for (let i = 0; i < redAssets.length; i++) {
      const redAsset = redAssets[i];
      console.log(`   ${i + 1}/${redAssets.length} 处理红色资产: ${redAsset['资产编码']} - ${redAsset['资产名称']}`);

      // 查找匹配的蓝色系统资产
      const matchingBlueAsset = this.findMatchingRedAsset(redAsset, blueAssets);

      if (matchingBlueAsset) {
        matchedCount++;
        usedBlueAssets.add(matchingBlueAsset['资产编码']);
        
        // 创建对照行（蓝色系统在左，红色系统在右）
        const comparisonRow = {
          // 蓝色系统字段
          '蓝色系统_资产编码': matchingBlueAsset['资产编码'],
          '蓝色系统_资产名称': matchingBlueAsset['资产名称'],
          '蓝色系统_资产等级': matchingBlueAsset['资产等级'],
          '蓝色系统_建筑面积': matchingBlueAsset['建筑面积'],
          '蓝色系统_租赁面积': matchingBlueAsset['租赁面积'],
          '蓝色系统_资产分类': matchingBlueAsset['资产分类'] || '',

          // 红色系统字段
          '红色系统_资产编码': redAsset['资产编码'],
          '红色系统_资产名称': redAsset['资产名称'],
          '红色系统_资产等级': redAsset['资产等级'],
          '红色系统_建筑面积': redAsset['建筑面积'],
          '红色系统_租赁面积': redAsset['租赁面积'],
          '红色系统_资产分类': redAsset['资产分类'] || '',

          // 匹配信息
          '匹配状态': '已匹配',
          '匹配方式': redAsset['OLD_AS_CODE'] === matchingBlueAsset['资产编码'] ? '编码' : '名称'
        };

        comparisonData.push(comparisonRow);
      } else {
        redUnmatchedCount++;
        // 创建红色系统未匹配行（红色系统在前，蓝色系统在后）
        const redUnmatchedRow = {
          // 红色系统字段
          '红色系统_资产编码': redAsset['资产编码'],
          '红色系统_资产名称': redAsset['资产名称'],
          '红色系统_资产等级': redAsset['资产等级'],
          '红色系统_建筑面积': redAsset['建筑面积'],
          '红色系统_租赁面积': redAsset['租赁面积'],
          '红色系统_资产分类': redAsset['资产分类'] || '',

          // 蓝色系统字段（空值）
          '蓝色系统_资产编码': '',
          '蓝色系统_资产名称': '',
          '蓝色系统_资产等级': '',
          '蓝色系统_建筑面积': '',
          '蓝色系统_租赁面积': '',
          '蓝色系统_资产分类': '',

          // 匹配信息
          '匹配状态': '未匹配',
          '匹配方式': ''
        };

        comparisonData.push(redUnmatchedRow);
      }
    }

    // 处理蓝色系统中未匹配的资产
    console.log('\n🔍 处理蓝色系统中未匹配的资产...');
    for (let i = 0; i < blueAssets.length; i++) {
      const blueAsset = blueAssets[i];
      
      // 跳过已经匹配的蓝色系统资产
      if (usedBlueAssets.has(blueAsset['资产编码'])) {
        continue;
      }

      console.log(`   ${i + 1}/${blueAssets.length} 处理蓝色资产: ${blueAsset['资产编码']} - ${blueAsset['资产名称']}`);
      
      blueUnmatchedCount++;
      // 创建蓝色系统未匹配行（左边有数据，右边空）
      const blueUnmatchedRow = {
        // 蓝色系统字段
        '蓝色系统_资产编码': blueAsset['资产编码'],
        '蓝色系统_资产名称': blueAsset['资产名称'],
        '蓝色系统_资产等级': blueAsset['资产等级'],
        '蓝色系统_建筑面积': blueAsset['建筑面积'],
        '蓝色系统_租赁面积': blueAsset['租赁面积'],

        // 红色系统字段（空值）
        '红色系统_资产编码': '',
        '红色系统_资产名称': '',
        '红色系统_资产等级': '',
        '红色系统_建筑面积': '',
          '红色系统_租赁面积': '',
          '红色系统_资产分类': '',

          // 匹配信息
        '匹配状态': '未匹配',
        '匹配方式': ''
      };

      comparisonData.push(blueUnmatchedRow);
    }

    console.log(`\n📊 匹配统计:`);
    console.log(`   已匹配: ${matchedCount} 条`);
    console.log(`   蓝色系统未匹配: ${blueUnmatchedCount} 条`);
    console.log(`   红色系统未匹配: ${redUnmatchedCount} 条`);
    console.log(`   总计: ${comparisonData.length} 条`);

    return comparisonData;
  }

  /**
   * 生成资产对照表Excel文件
   * @param {Array} comparisonData - 对照表数据
   * @param {string} sheetName - 工作表名称
   * @returns {string} 生成的文件路径
   */
  generateComparisonExcel(comparisonData, sheetName = '资产对照表') {
    try {
      // 确保输出目录存在
      const outputDir = this.config.excel.outputDir;
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
        console.log(`✅ 创建输出目录: ${outputDir}`);
      }

      // 创建工作簿
      const workbook = XLSX.utils.book_new();

      // 准备表头（根据对照表类型调整字段顺序）
      let headers;
      if (sheetName.includes('红色系统')) {
        // 红色系统对照表：红色系统字段在前
        headers = [
          // 红色系统字段
          '资产编码', '资产名称', '资产等级', '建筑面积', '租赁面积', '资产分类',
          // 蓝色系统字段
          '资产编码', '资产名称', '资产等级', '建筑面积', '租赁面积', '资产分类',
          // 匹配信息
          '匹配状态', '匹配方式'
        ];
      } else {
        // 蓝色系统对照表：蓝色系统字段在前
        headers = [
          // 蓝色系统字段
          '资产编码', '资产名称', '资产等级', '建筑面积', '租赁面积', '资产分类',
          // 红色系统字段
          '资产编码', '资产名称', '资产等级', '建筑面积', '租赁面积', '资产分类',
          // 匹配信息
          '匹配状态', '匹配方式'
        ];
      }

      // 计算匹配统计
      const matchedCount = comparisonData.filter(item => item['匹配状态'] === '已匹配').length;
      const unmatchedCount = comparisonData.filter(item => item['匹配状态'] === '未匹配').length;
      const totalCount = comparisonData.length;

      // 将数据转换为二维数组格式（根据对照表类型调整字段顺序）
      const dataRows = comparisonData.map(row => {
        if (sheetName.includes('红色系统')) {
          // 红色系统对照表：红色系统字段在前
          return [
            row['红色系统_资产编码'],
            row['红色系统_资产名称'],
            row['红色系统_资产等级'],
            row['红色系统_建筑面积'],
            row['红色系统_租赁面积'],
            row['红色系统_资产分类'],
            row['蓝色系统_资产编码'],
            row['蓝色系统_资产名称'],
            row['蓝色系统_资产等级'],
            row['蓝色系统_建筑面积'],
            row['蓝色系统_租赁面积'],
            row['蓝色系统_资产分类'],
            row['匹配状态'],
            row['匹配方式']
          ];
        } else {
          // 蓝色系统对照表：蓝色系统字段在前
          return [
            row['蓝色系统_资产编码'],
            row['蓝色系统_资产名称'],
            row['蓝色系统_资产等级'],
            row['蓝色系统_建筑面积'],
            row['蓝色系统_租赁面积'],
            row['蓝色系统_资产分类'],
            row['红色系统_资产编码'],
            row['红色系统_资产名称'],
            row['红色系统_资产等级'],
            row['红色系统_建筑面积'],
            row['红色系统_租赁面积'],
            row['红色系统_资产分类'],
            row['匹配状态'],
            row['匹配方式']
          ];
        }
      });

      const worksheetData = [headers, ...dataRows];

      // 创建工作表
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

      // 设置列宽
      const columnWidths = [
        // 蓝色系统列宽
        { wch: 20 }, { wch: 25 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 15 },
        // 红色系统列宽
        { wch: 20 }, { wch: 25 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 15 },
        // 匹配信息列宽
        { wch: 10 }, { wch: 15 }
      ];
      worksheet['!cols'] = columnWidths;

      // 设置表头样式（兼容WPS）
      const headerRange = XLSX.utils.decode_range(worksheet['!ref']);
      
      // 表头行设置粗体
      for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!worksheet[cellAddress]) worksheet[cellAddress] = { v: '' };
        worksheet[cellAddress].s = {
          font: { bold: true }
        };
      }

      // 设置行高
      worksheet['!rows'] = [
        { hpt: 20 }, // 表头行高
        ...comparisonData.map(() => ({ hpt: 15 })) // 数据行高
      ];

      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      // 生成文件路径
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const managementAreaName = this.config.managementArea.name;
      const filename = `${sheetName}_${managementAreaName}_${timestamp}.xlsx`;
      const filePath = path.join(outputDir, filename);

      // 写入文件
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ 资产对照表生成成功: ${filePath}`);
      console.log(`📊 共生成 ${comparisonData.length} 条对照记录`);

      return filePath;
    } catch (error) {
      console.error('❌ 资产对照表生成失败:', error.message);
      throw error;
    }
  }
}

module.exports = AssetComparisonGenerator;
