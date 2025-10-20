const XLSX = require('xlsx');
const mysql = require('mysql2/promise');
const config = require('./config');
const { v4: uuidv4 } = require('crypto');

/**
 * 资产编码对照表导入器
 * 用于将Excel文件中的老资产编码和新资产编码对照关系导入到数据库
 */
class AssetCodeImporter {
  constructor() {
    this.connection = null;
  }

  /**
   * 建立数据库连接
   */
  async connect() {
    try {
      const dbConfig = config.databases[config.currentDatabase];
      if (!dbConfig) {
        throw new Error(`数据库配置 "${config.currentDatabase}" 不存在`);
      }
      
      this.connection = await mysql.createConnection(dbConfig);
      console.log(`✅ 数据库连接成功 (${dbConfig.name})`);
      return this.connection;
    } catch (error) {
      console.error('❌ 数据库连接失败:', error.message);
      throw error;
    }
  }

  /**
   * 读取Excel文件
   * @param {string} filePath - Excel文件路径
   * @returns {Array} 解析后的数据数组
   */
  readExcelFile(filePath) {
    try {
      console.log(`📖 正在读取Excel文件: ${filePath}`);
      
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // 将工作表转换为JSON数组，第一行作为标题
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (data.length < 2) {
        throw new Error('Excel文件数据不足，至少需要包含标题行和一行数据');
      }

      // 获取标题行
      const headers = data[0];
      console.log('📋 Excel文件标题:', headers);

      // 验证必要的列是否存在
      const oldCodeIndex = headers.findIndex(h => h && h.toString().trim().toLowerCase().includes('old'));
      const newCodeIndex = headers.findIndex(h => h && h.toString().trim().toLowerCase().includes('new'));

      if (oldCodeIndex === -1 || newCodeIndex === -1) {
        throw new Error('Excel文件必须包含OLD_AS_CODE和NEW_AS_CODE列');
      }

      // 处理数据行，处理重复的OLD_AS_CODE
      const processedData = [];
      const duplicateMap = new Map(); // 用于跟踪重复数据
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row && row.length > 0 && row[oldCodeIndex] && row[newCodeIndex]) {
          const oldAsCode = row[oldCodeIndex].toString().trim();
          const newAsCode = row[newCodeIndex].toString().trim();
          
          if (duplicateMap.has(oldAsCode)) {
            // 发现重复的OLD_AS_CODE
            const existingEntry = duplicateMap.get(oldAsCode);
            if (existingEntry.newAsCode !== newAsCode) {
              console.warn(`⚠️  发现冲突: OLD_AS_CODE "${oldAsCode}" 对应不同的NEW_AS_CODE`);
              console.warn(`   第${existingEntry.row}行: ${oldAsCode} -> ${existingEntry.newAsCode}`);
              console.warn(`   第${i}行: ${oldAsCode} -> ${newAsCode}`);
              console.warn(`   将使用第${i}行的数据（后出现的为准）`);
            } else {
              console.log(`ℹ️  发现重复: OLD_AS_CODE "${oldAsCode}" 对应相同的NEW_AS_CODE "${newAsCode}"`);
            }
          }
          
          // 记录或更新数据
          duplicateMap.set(oldAsCode, {
            oldAsCode,
            newAsCode,
            row: i
          });
        }
      }

      // 将Map转换为数组
      processedData.push(...duplicateMap.values());

      console.log(`✅ 成功读取 ${processedData.length} 条唯一数据记录`);
      console.log(`📊 原始数据行数: ${data.length - 1}, 去重后: ${processedData.length}`);
      
      return processedData;
    } catch (error) {
      console.error('❌ 读取Excel文件失败:', error.message);
      throw error;
    }
  }

  /**
   * 创建数据库表（如果不存在）
   */
  async createTableIfNotExists() {
    const createTableSQL = `
      CREATE TABLE IF NOT EXISTS \`old_as_code_new\` (
        \`ID\` varchar(32) NOT NULL COMMENT '用户ID',
        \`OLD_AS_CODE\` varchar(50) NOT NULL COMMENT '老资产编码',
        \`NEW_AS_CODE\` varchar(80) NOT NULL COMMENT '新资产编码',
        PRIMARY KEY (\`ID\`),
        UNIQUE KEY \`unique_old_as_code\` (\`OLD_AS_CODE\`)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC COMMENT='老资产对照表'
    `;

    try {
      await this.connection.execute(createTableSQL);
      console.log('✅ 数据库表创建成功或已存在');
    } catch (error) {
      console.error('❌ 创建数据库表失败:', error.message);
      throw error;
    }
  }

  /**
   * 更新现有表的字段长度
   */
  async updateTableStructure() {
    try {
      console.log('🔧 检查并更新表结构...');
      
      // 检查表是否存在
      const [tables] = await this.connection.execute(
        "SHOW TABLES LIKE 'old_as_code_new'"
      );
      
      if (tables.length > 0) {
        // 表存在，检查字段长度
        const [columns] = await this.connection.execute(
          "SHOW COLUMNS FROM `old_as_code_new` WHERE Field = 'OLD_AS_CODE'"
        );
        
        if (columns.length > 0) {
          const column = columns[0];
          const currentLength = parseInt(column.Type.match(/varchar\((\d+)\)/)?.[1] || '0');
          
          if (currentLength < 50) {
            console.log(`📏 当前OLD_AS_CODE字段长度: ${currentLength}, 需要更新为50`);
            await this.connection.execute(
              "ALTER TABLE `old_as_code_new` MODIFY COLUMN `OLD_AS_CODE` varchar(50) NOT NULL COMMENT '老资产编码'"
            );
            console.log('✅ OLD_AS_CODE字段长度已更新为50');
          } else {
            console.log('✅ OLD_AS_CODE字段长度已足够');
          }
        }
      }
    } catch (error) {
      console.error('❌ 更新表结构失败:', error.message);
      throw error;
    }
  }

  /**
   * 清空现有数据
   */
  async clearExistingData() {
    try {
      await this.connection.execute('DELETE FROM `old_as_code_new`');
      console.log('✅ 已清空现有数据');
    } catch (error) {
      console.error('❌ 清空数据失败:', error.message);
      throw error;
    }
  }

  /**
   * 批量插入数据
   * @param {Array} data - 要插入的数据数组
   */
  async insertData(data) {
    if (!data || data.length === 0) {
      console.log('⚠️ 没有数据需要插入');
      return;
    }

    try {
      console.log(`📝 开始插入 ${data.length} 条数据...`);
      
      // 准备批量插入的SQL语句
      const insertSQL = `
        INSERT INTO \`old_as_code_new\` (\`ID\`, \`OLD_AS_CODE\`, \`NEW_AS_CODE\`) 
        VALUES (?, ?, ?)
        ON DUPLICATE KEY UPDATE 
        \`NEW_AS_CODE\` = VALUES(\`NEW_AS_CODE\`)
      `;

      let insertedCount = 0;
      let updatedCount = 0;

      // 逐条插入以便跟踪插入和更新情况
      for (const item of data) {
        try {
          // 先检查是否已存在
          const [existingRows] = await this.connection.execute(
            'SELECT NEW_AS_CODE FROM `old_as_code_new` WHERE OLD_AS_CODE = ?',
            [item.oldAsCode]
          );

          const value = [
            this.generateId(), // 生成唯一ID
            item.oldAsCode,
            item.newAsCode
          ];

          await this.connection.execute(insertSQL, value);

          if (existingRows.length > 0) {
            updatedCount++;
            console.log(`🔄 更新: ${item.oldAsCode} -> ${item.newAsCode}`);
          } else {
            insertedCount++;
            console.log(`➕ 新增: ${item.oldAsCode} -> ${item.newAsCode}`);
          }
        } catch (error) {
          console.error(`❌ 插入数据失败 (${item.oldAsCode}):`, error.message);
          throw error;
        }
      }

      console.log(`✅ 数据插入完成: 新增 ${insertedCount} 条, 更新 ${updatedCount} 条`);
    } catch (error) {
      console.error('❌ 插入数据失败:', error.message);
      throw error;
    }
  }

  /**
   * 生成唯一ID
   * @returns {string} 唯一ID
   */
  generateId() {
    return require('crypto').randomBytes(16).toString('hex');
  }

  /**
   * 验证导入的数据
   */
  async validateImportedData() {
    try {
      const [rows] = await this.connection.execute('SELECT COUNT(*) as count FROM `old_as_code_new`');
      const count = rows[0].count;
      console.log(`✅ 验证完成，数据库中共有 ${count} 条记录`);
      
      // 显示前几条记录作为示例
      const [sampleRows] = await this.connection.execute(
        'SELECT OLD_AS_CODE, NEW_AS_CODE FROM `old_as_code_new` LIMIT 5'
      );
      
      if (sampleRows.length > 0) {
        console.log('📋 数据示例:');
        sampleRows.forEach((row, index) => {
          console.log(`  ${index + 1}. ${row.OLD_AS_CODE} -> ${row.NEW_AS_CODE}`);
        });
      }
    } catch (error) {
      console.error('❌ 验证数据失败:', error.message);
      throw error;
    }
  }

  /**
   * 执行完整的导入流程
   * @param {string} excelFilePath - Excel文件路径
   * @param {boolean} clearExisting - 是否清空现有数据
   */
  async importFromExcel(excelFilePath, clearExisting = true) {
    try {
      console.log('🚀 开始资产编码对照表导入流程...');
      
      // 1. 连接数据库
      await this.connect();
      
      // 2. 创建表（如果不存在）
      await this.createTableIfNotExists();
      
      // 3. 更新表结构（确保字段长度足够）
      await this.updateTableStructure();
      
      // 4. 清空现有数据（如果需要）
      if (clearExisting) {
        await this.clearExistingData();
      }
      
      // 5. 读取Excel文件
      const data = this.readExcelFile(excelFilePath);
      
      // 6. 插入数据
      await this.insertData(data);
      
      // 7. 验证导入结果
      await this.validateImportedData();
      
      console.log('🎉 资产编码对照表导入完成！');
      
    } catch (error) {
      console.error('❌ 导入过程中发生错误:', error.message);
      throw error;
    } finally {
      // 关闭数据库连接
      if (this.connection) {
        await this.connection.end();
        console.log('✅ 数据库连接已关闭');
      }
    }
  }

  /**
   * 关闭数据库连接
   */
  async close() {
    if (this.connection) {
      await this.connection.end();
      console.log('✅ 数据库连接已关闭');
    }
  }
}

module.exports = AssetCodeImporter;
