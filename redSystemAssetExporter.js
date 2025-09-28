const DatabaseManager = require('./database');
const ExcelGenerator = require('./excelGenerator');

class RedSystemAssetExporter {
  constructor(config) {
    this.dbManager = new DatabaseManager(config, 'redSystem');
    this.excelGenerator = new ExcelGenerator(config);
    this.config = config;
  }

  /**
   * 查询顶级资产（-99级）
   * @returns {Array} 顶级资产列表
   */
  async getTopLevelAssets() {
    const managementAreaName = this.config.managementArea.name;
    const query = `
      SELECT
        AS_CODE AS 资产编码,
        AS_NAME AS 资产名称,
        AS_LV AS 资产等级,
        OPERATING AS 资产类型,
        AS_TYPE_NAME AS 资产分类,
        AS_ADDRESS AS 资产地址,
        COALESCE(AS_CONSTRUCTION_AREA, 0) AS 建筑面积,
        COALESCE(AS_USABLE_AREA, 0) AS 租赁面积,
        UP_AS_CODE AS 上级资产编码,
        NEW_AS_CODE,
        NEW_AS_NAME,
        OLD_AS_CODE,
        OLD_AS_NAME,
        AS_STATE,
        U_DELETE 
      FROM
        as_asset 
      WHERE
        OPERATING_NAME = ? 
        AND U_DELETE = 1 
        AND AS_LV = -99
        AND AS_STATE IN ('CHECKED', 'INIT')
      ORDER BY
        AS_CODE
    `;

    return await this.dbManager.query(query, [managementAreaName]);
  }

  /**
   * 查询指定父级资产的所有子资产
   * @param {string} parentCode - 父级资产编码
   * @returns {Array} 子资产列表
   */
  async getChildAssets(parentCode) {
    const managementAreaName = this.config.managementArea.name;
    const query = `
      SELECT
        AS_CODE AS 资产编码,
        AS_NAME AS 资产名称,
        AS_LV AS 资产等级,
        OPERATING AS 资产类型,
        AS_TYPE_NAME AS 资产分类,
        AS_ADDRESS AS 资产地址,
        COALESCE(AS_CONSTRUCTION_AREA, 0) AS 建筑面积,
        COALESCE(AS_USABLE_AREA, 0) AS 租赁面积,
        UP_AS_CODE AS 上级资产编码,
        NEW_AS_CODE,
        NEW_AS_NAME,
        OLD_AS_CODE,
        OLD_AS_NAME,
        AS_STATE,
        U_DELETE 
      FROM
        as_asset 
      WHERE
        OPERATING_NAME = ? 
        AND U_DELETE = 1 
        AND AS_STATE IN ('CHECKED', 'INIT')
        AND UP_AS_CODE = ?
      ORDER BY
        AS_CODE
    `;

    return await this.dbManager.query(query, [managementAreaName, parentCode]);
  }

  /**
   * 递归查询所有层级的子资产
   * @param {string} parentCode - 父级资产编码
   * @param {number} currentLevel - 当前层级
   * @returns {Array} 所有子资产列表
   */
  async getAllChildAssets(parentCode, currentLevel = 1) {
    const directChildren = await this.getChildAssets(parentCode);
    let allChildren = [...directChildren];

    // 如果不是3级资产，继续查询下一级
    if (currentLevel < 3) {
      for (const child of directChildren) {
        const grandChildren = await this.getAllChildAssets(child['资产编码'], currentLevel + 1);
        allChildren = allChildren.concat(grandChildren);
      }
    }

    return allChildren;
  }

  /**
   * 构建层级资产汇总数据
   * @returns {Array} 包含父子关系的汇总数据
   */
  async buildHierarchicalData() {
    try {
      const managementAreaName = this.config.managementArea.name;
      const dbName = this.config.databases.redSystem.name;
      console.log(`🔍 正在查询数据库 "${dbName}" 中管理片区 "${managementAreaName}" 的顶级资产（-99级）...`);
      
      const topLevelAssets = await this.getTopLevelAssets();
      
      if (topLevelAssets.length === 0) {
        console.log('⚠️  未找到顶级资产数据');
        return [];
      }

      console.log(`📊 找到 ${topLevelAssets.length} 个顶级资产，开始查询所有子资产...`);
      
      const hierarchicalData = [];

      for (let i = 0; i < topLevelAssets.length; i++) {
        const parentAsset = topLevelAssets[i];
        console.log(`   ${i + 1}/${topLevelAssets.length} 处理资产: ${parentAsset['资产编码']} - ${parentAsset['资产名称']}`);

        // 递归查询该顶级资产的所有子资产
        const allChildAssets = await this.getAllChildAssets(parentAsset['资产编码']);
        
        // 构建子资产编码列表（只显示直接下级）
        const directChildren = await this.getChildAssets(parentAsset['资产编码']);
        const childCodes = directChildren.map(child => child['资产编码']).join('\n');
        
        // 添加顶级资产记录，包含下级资产编码列表
        const parentRecord = {
          ...parentAsset,
          '下级资产编码列表': childCodes || '',
          '上级资产编码': '' // 顶级资产没有上级
        };
        hierarchicalData.push(parentRecord);

        // 添加所有子资产记录，包含上级资产编码和下级资产编码列表
        for (const childAsset of allChildAssets) {
          // 获取该子资产的直接下级资产
          const childDirectChildren = await this.getChildAssets(childAsset['资产编码']);
          const childCodes = childDirectChildren.map(child => child['资产编码']).join('\n');
          
          // 调试信息：显示有下级资产的记录
          if (childDirectChildren.length > 0) {
            console.log(`     📋 ${childAsset['资产编码']} (${childAsset['资产等级']}级) 有 ${childDirectChildren.length} 个下级资产: ${childCodes}`);
          }
          
          const childRecord = {
            ...childAsset,
            '下级资产编码列表': childCodes || '', // 根据实际情况设置下级资产编码列表
            '上级资产编码': this.getParentCode(childAsset, allChildAssets, topLevelAssets)
          };
          hierarchicalData.push(childRecord);
        }

        console.log(`     ✅ 找到 ${allChildAssets.length} 个子资产（包含所有层级）`);
      }

      console.log(`📈 层级数据构建完成，总计 ${hierarchicalData.length} 条记录`);
      return hierarchicalData;

    } catch (error) {
      console.error('❌ 构建层级数据失败:', error.message);
      throw error;
    }
  }

  /**
   * 获取资产的上级资产编码
   * @param {Object} asset - 当前资产
   * @param {Array} allChildren - 所有子资产列表
   * @param {Array} topAssets - 顶级资产列表
   * @returns {string} 上级资产编码
   */
  getParentCode(asset, allChildren, topAssets) {
    // 如果UP_AS_CODE在顶级资产中，直接返回
    const topParent = topAssets.find(top => top['资产编码'] === asset['上级资产编码']);
    if (topParent) {
      return asset['上级资产编码'];
    }

    // 否则在子资产中查找
    const parent = allChildren.find(child => child['资产编码'] === asset['上级资产编码']);
    return parent ? asset['上级资产编码'] : '';
  }

  /**
   * 生成层级资产Excel文件
   * @returns {string} 生成的文件路径
   */
  async exportToExcel() {
    try {
      const managementAreaName = this.config.managementArea.name;
      const dbName = this.config.databases.redSystem.name;
      console.log(`🚀 开始导出数据库 "${dbName}" 中管理片区 "${managementAreaName}" 的层级资产数据...`);
      
      // 连接数据库
      await this.dbManager.connect();

      // 构建层级数据
      const hierarchicalData = await this.buildHierarchicalData();

      if (hierarchicalData.length === 0) {
        console.log('⚠️  没有数据可导出');
        return null;
      }

      // 生成Excel文件
      console.log('📊 正在生成Excel文件...');
      const filePath = this.excelGenerator.generateRedSystemExcel(
        hierarchicalData, 
        '红色系统资产层级数据'
      );

      // 显示统计信息
      const topLevelCount = hierarchicalData.filter(item => item['资产等级'] === -99).length;
      const level1Count = hierarchicalData.filter(item => item['资产等级'] === 1).length;
      const level2Count = hierarchicalData.filter(item => item['资产等级'] === 2).length;
      const level3Count = hierarchicalData.filter(item => item['资产等级'] === 3).length;

      console.log('\n📈 导出完成统计:');
      console.log(`   顶级资产(-99级): ${topLevelCount} 个`);
      console.log(`   1级资产: ${level1Count} 个`);
      console.log(`   2级资产: ${level2Count} 个`);
      console.log(`   3级资产: ${level3Count} 个`);
      console.log(`   总记录数: ${hierarchicalData.length}`);
      console.log(`   文件路径: ${filePath}`);

      return filePath;

    } catch (error) {
      console.error('❌ 导出失败:', error.message);
      throw error;
    } finally {
      // 关闭数据库连接
      await this.dbManager.close();
    }
  }
}

module.exports = RedSystemAssetExporter;
