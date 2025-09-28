const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class ExcelGenerator {
  constructor(config) {
    this.config = config;
  }

  /**
   * 确保输出目录存在
   */
  ensureOutputDir() {
    const outputDir = this.config.excel.outputDir;
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`✅ 创建输出目录: ${outputDir}`);
    }
  }

  /**
   * 生成Excel文件
   * @param {Array} data - 要导出的数据
   * @param {string} sheetName - 工作表名称
   * @returns {string} 生成的文件路径
   */
  generateExcel(data, sheetName = '资产数据') {
    try {
      this.ensureOutputDir();

      // 创建工作簿
      const workbook = XLSX.utils.book_new();

      // 将数据转换为工作表
      const worksheet = XLSX.utils.json_to_sheet(data);

      // 设置列宽
      const columnWidths = [
        { wch: 15 }, // 资产编码
        { wch: 25 }, // 资产名称
        { wch: 10 }, // 资产等级
        { wch: 15 }, // 资产类型
        { wch: 20 }, // 资产分类
        { wch: 30 }, // 资产地址
        { wch: 12 }, // 建筑面积
        { wch: 12 }, // 租赁面积
        { wch: 15 }, // 上级资产编码
        { wch: 10 }, // AS_STATE
        { wch: 10 }  // U_DELETE
      ];
      worksheet['!cols'] = columnWidths;

      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      // 生成文件路径
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const filename = `${this.config.excel.filename.replace('.xlsx', '')}_${timestamp}.xlsx`;
      const filePath = path.join(this.config.excel.outputDir, filename);

      // 写入文件
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ Excel文件生成成功: ${filePath}`);
      console.log(`📊 共导出 ${data.length} 条记录`);

      return filePath;
    } catch (error) {
      console.error('❌ Excel文件生成失败:', error.message);
      throw error;
    }
  }

  /**
   * 生成层级资产Excel文件
   * @param {Array} data - 要导出的层级数据
   * @param {string} sheetName - 工作表名称
   * @returns {string} 生成的文件路径
   */
  generateHierarchicalExcel(data, sheetName = '层级资产数据') {
    try {
      this.ensureOutputDir();

      // 创建工作簿
      const workbook = XLSX.utils.book_new();

      // 准备数据，添加表头样式
      const headers = [
        '资产编码', '合同编号', '资产名称', '资产等级', '资产类型', '资产分类',
        '资产地址', '建筑面积', '租赁面积', '上级资产编码', '下级资产编码列表',
        'NEW_AS_CODE', 'NEW_AS_NAME', 'OLD_AS_CODE', 'OLD_AS_NAME'
      ];

      // 将数据转换为二维数组格式
      const worksheetData = [headers, ...data.map(row => [
        row['资产编码'],
        row['合同编号'] || '',
        row['资产名称'],
        row['资产等级'],
        row['资产类型'],
        row['资产分类'],
        row['资产地址'],
        row['建筑面积'],
        row['租赁面积'],
        row['上级资产编码'],
        row['下级资产编码列表'],
        row['NEW_AS_CODE'] || '',
        row['NEW_AS_NAME'] || '',
        row['OLD_AS_CODE'] || '',
        row['OLD_AS_NAME'] || ''
      ])];

      // 创建工作表
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

      // 设置列宽
      const columnWidths = [
        { wch: 20 }, // 资产编码
        { wch: 20 }, // 合同编号
        { wch: 25 }, // 资产名称
        { wch: 10 }, // 资产等级
        { wch: 15 }, // 资产类型
        { wch: 25 }, // 资产分类
        { wch: 30 }, // 资产地址
        { wch: 12 }, // 建筑面积
        { wch: 12 }, // 租赁面积
        { wch: 20 }, // 上级资产编码
        { wch: 40 }, // 下级资产编码列表
        { wch: 20 }, // NEW_AS_CODE
        { wch: 25 }, // NEW_AS_NAME
        { wch: 20 }, // OLD_AS_CODE
        { wch: 25 }  // OLD_AS_NAME
      ];
      worksheet['!cols'] = columnWidths;

      // 设置行高，根据下级资产编码列表的内容调整
      worksheet['!rows'] = [
        { hpt: 20 }, // 表头行高
        ...data.map(row => {
          // 如果下级资产编码列表有内容，根据换行数量调整行高
          const childCodes = row['下级资产编码列表'] || '';
          const lineCount = childCodes.split('\n').length;
          return { hpt: Math.max(15, lineCount * 15) }; // 每行至少15pt，多行时增加高度
        })
      ];

      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      // 生成文件路径
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const managementAreaName = this.config.managementArea.name;
      const filename = `蓝色系统_${managementAreaName}_资产数据_${timestamp}.xlsx`;
      const filePath = path.join(this.config.excel.outputDir, filename);

      // 写入文件
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ 层级Excel文件生成成功: ${filePath}`);
      console.log(`📊 共导出 ${data.length} 条记录`);

      return filePath;
    } catch (error) {
      console.error('❌ 层级Excel文件生成失败:', error.message);
      throw error;
    }
  }

  /**
   * 生成红色系统层级资产Excel文件
   * @param {Array} data - 要导出的层级数据
   * @param {string} sheetName - 工作表名称
   * @returns {string} 生成的文件路径
   */
  generateRedSystemExcel(data, sheetName = '红色系统层级资产数据') {
    try {
      this.ensureOutputDir();

      // 创建工作簿
      const workbook = XLSX.utils.book_new();

      // 准备数据，添加表头样式（包含红色系统特有字段）
      const headers = [
        '资产编码', '合同编号', '资产名称', '资产等级', '资产类型', '资产分类',
        '资产地址', '建筑面积', '租赁面积', '上级资产编码', '下级资产编码列表',
        'NEW_AS_CODE', 'NEW_AS_NAME', 'OLD_AS_CODE', 'OLD_AS_NAME'
      ];

      // 将数据转换为二维数组格式
      const worksheetData = [headers, ...data.map(row => [
        row['资产编码'],
        row['合同编号'] || '',
        row['资产名称'],
        row['资产等级'],
        row['资产类型'],
        row['资产分类'],
        row['资产地址'],
        row['建筑面积'],
        row['租赁面积'],
        row['上级资产编码'],
        row['下级资产编码列表'],
        row['NEW_AS_CODE'] || '',
        row['NEW_AS_NAME'] || '',
        row['OLD_AS_CODE'] || '',
        row['OLD_AS_NAME'] || ''
      ])];

      // 创建工作表
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

      // 设置列宽
      const columnWidths = [
        { wch: 20 }, // 资产编码
        { wch: 20 }, // 合同编号
        { wch: 25 }, // 资产名称
        { wch: 10 }, // 资产等级
        { wch: 15 }, // 资产类型
        { wch: 25 }, // 资产分类
        { wch: 30 }, // 资产地址
        { wch: 12 }, // 建筑面积
        { wch: 12 }, // 租赁面积
        { wch: 20 }, // 上级资产编码
        { wch: 40 }, // 下级资产编码列表
        { wch: 20 }, // NEW_AS_CODE
        { wch: 25 }, // NEW_AS_NAME
        { wch: 20 }, // OLD_AS_CODE
        { wch: 25 }  // OLD_AS_NAME
      ];
      worksheet['!cols'] = columnWidths;

      // 设置行高，根据下级资产编码列表的内容调整
      worksheet['!rows'] = [
        { hpt: 20 }, // 表头行高
        ...data.map(row => {
          // 如果下级资产编码列表有内容，根据换行数量调整行高
          const childCodes = row['下级资产编码列表'] || '';
          const lineCount = childCodes.split('\n').length;
          return { hpt: Math.max(15, lineCount * 15) }; // 每行至少15pt，多行时增加高度
        })
      ];

      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      // 生成文件路径
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const managementAreaName = this.config.managementArea.name;
      const filename = `红色系统_${managementAreaName}_层级资产数据_${timestamp}.xlsx`;
      const filePath = path.join(this.config.excel.outputDir, filename);

      // 写入文件
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ 红色系统Excel文件生成成功: ${filePath}`);
      console.log(`📊 共导出 ${data.length} 条记录`);

      return filePath;
    } catch (error) {
      console.error('❌ 红色系统Excel文件生成失败:', error.message);
      throw error;
    }
  }


  /**
   * 生成带样式的Excel文件（高级版本）
   * @param {Array} data - 要导出的数据
   * @param {string} sheetName - 工作表名称
   * @returns {string} 生成的文件路径
   */
  generateStyledExcel(data, sheetName = '资产数据') {
    try {
      this.ensureOutputDir();

      // 创建工作簿
      const workbook = XLSX.utils.book_new();

      // 准备数据，添加表头样式
      const headers = [
        '资产编码', '资产名称', '资产等级', '资产类型', '资产分类',
        '资产地址', '建筑面积', '租赁面积', '上级资产编码', 'AS_STATE', 'U_DELETE'
      ];

      // 将数据转换为二维数组格式
      const worksheetData = [headers, ...data.map(row => [
        row['资产编码'],
        row['资产名称'],
        row['资产等级'],
        row['资产类型'],
        row['资产分类'],
        row['资产地址'],
        row['建筑面积'],
        row['租赁面积'],
        row['上级资产编码'],
        row['AS_STATE'],
        row['U_DELETE']
      ])];

      // 创建工作表
      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

      // 设置列宽
      const columnWidths = [
        { wch: 15 }, // 资产编码
        { wch: 25 }, // 资产名称
        { wch: 10 }, // 资产等级
        { wch: 15 }, // 资产类型
        { wch: 20 }, // 资产分类
        { wch: 30 }, // 资产地址
        { wch: 12 }, // 建筑面积
        { wch: 12 }, // 租赁面积
        { wch: 15 }, // 上级资产编码
        { wch: 10 }, // AS_STATE
        { wch: 10 }  // U_DELETE
      ];
      worksheet['!cols'] = columnWidths;

      // 设置行高
      worksheet['!rows'] = [
        { hpt: 20 }, // 表头行高
        ...data.map(() => ({ hpt: 15 })) // 数据行高
      ];

      // 添加工作表到工作簿
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      // 生成文件路径
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const filename = `${this.config.excel.filename.replace('.xlsx', '')}_${timestamp}.xlsx`;
      const filePath = path.join(this.config.excel.outputDir, filename);

      // 写入文件
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ 带样式的Excel文件生成成功: ${filePath}`);
      console.log(`📊 共导出 ${data.length} 条记录`);

      return filePath;
    } catch (error) {
      console.error('❌ 带样式的Excel文件生成失败:', error.message);
      throw error;
    }
  }
}

module.exports = ExcelGenerator;
