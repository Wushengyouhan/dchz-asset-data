const DatabaseManager = require('./database');
const ExcelGenerator = require('./excelGenerator');

class HierarchicalAssetExporter {
  constructor(config, databaseKey = null) {
    // 如果没有指定数据库，使用配置中的当前数据库
    const dbKey = databaseKey || config.currentDatabase;
    this.dbManager = new DatabaseManager(config, dbKey);
    this.excelGenerator = new ExcelGenerator(config);
    this.config = config;
    this.databaseKey = dbKey;
  }

  /**
   * 查询一级资产列表
   * @returns {Array} 一级资产列表
   */
  async getLevel1Assets() {
    const managementAreaName = this.config.managementArea.name;
    const query = `
      SELECT DISTINCT
        a.AS_CODE AS 资产编码,
        a.AS_NAME AS 资产名称,
        a.AS_LV AS 资产等级,
        a.OPERATING AS 资产类型,
        a.AS_TYPE_NAME AS 资产分类,
        a.AS_ADDRESS AS 资产地址,
        COALESCE(a.AS_CONSTRUCTION_AREA, 0) AS 建筑面积,
        COALESCE(a.AS_USABLE_AREA, 0) AS 租赁面积,
        a.UP_AS_CODE AS 上级资产编码,
        a.AS_STATE,
        a.U_DELETE,
        a.NEW_AS_CODE,
        a.NEW_AS_NAME,
        a.OLD_AS_CODE,
        a.OLD_AS_NAME,
        c.CON_CODE AS 合同编号
      FROM
        as_asset a
      LEFT JOIN (
        SELECT
          ccd.AS_CODE,
          c.CON_CODE
        FROM con_contracts_detail ccd
        INNER JOIN con_contracts c ON ccd.CONTRACTS_ID = c.ID
        WHERE c.U_DELETE = 1
          AND c.START_DATE < NOW()
          AND c.END_DATE > NOW()
          AND c.CON_STATE IN ('CHECKED', 'INIT', 'WORKFLOWED')
      ) c ON a.AS_CODE = c.AS_CODE
      WHERE
        a.OPERATING_NAME = ? 
        AND a.U_DELETE = 1 
        AND a.AS_STATE LIKE '%BLUE' 
        AND a.AS_LV = 1 
      ORDER BY
        a.AS_CODE
    `;

    return await this.dbManager.query(query, [managementAreaName]);
  }

  /**
   * 查询指定一级资产的所有子资产
   * @param {string} parentCode - 父级资产编码
   * @returns {Array} 子资产列表
   */
  async getChildAssets(parentCode) {
    const managementAreaName = this.config.managementArea.name;
    const query = `
      SELECT DISTINCT
        a.AS_CODE AS 资产编码,
        a.AS_NAME AS 资产名称,
        a.AS_LV AS 资产等级,
        a.OPERATING AS 资产类型,
        a.AS_TYPE_NAME AS 资产分类,
        a.AS_ADDRESS AS 资产地址,
        COALESCE(a.AS_CONSTRUCTION_AREA, 0) AS 建筑面积,
        COALESCE(a.AS_USABLE_AREA, 0) AS 租赁面积,
        a.UP_AS_CODE AS 上级资产编码,
        a.AS_STATE,
        a.U_DELETE,
        a.NEW_AS_CODE,
        a.NEW_AS_NAME,
        a.OLD_AS_CODE,
        a.OLD_AS_NAME,
        c.CON_CODE AS 合同编号
      FROM
        as_asset a
      LEFT JOIN (
        SELECT
          ccd.AS_CODE,
          c.CON_CODE
        FROM con_contracts_detail ccd
        INNER JOIN con_contracts c ON ccd.CONTRACTS_ID = c.ID
        WHERE c.U_DELETE = 1
          AND c.START_DATE < NOW()
          AND c.END_DATE > NOW()
          AND c.CON_STATE IN ('CHECKED', 'INIT', 'WORKFLOWED')
      ) c ON a.AS_CODE = c.AS_CODE
      WHERE
        a.OPERATING_NAME = ? 
        AND a.U_DELETE = 1 
        AND a.AS_STATE LIKE '%BLUE' 
        AND a.UP_AS_CODE = ?
      ORDER BY
        a.AS_CODE
    `;

    return await this.dbManager.query(query, [managementAreaName, parentCode]);
  }

  /**
   * 构建层级资产汇总数据
   * @returns {Array} 包含父子关系的汇总数据
   */
  async buildHierarchicalData() {
    try {
      const managementAreaName = this.config.managementArea.name;
      const dbName = this.config.databases[this.databaseKey].name;
      console.log(`🔍 正在查询数据库 "${dbName}" 中管理片区 "${managementAreaName}" 的一级资产...`);
      const level1Assets = await this.getLevel1Assets();
      
      if (level1Assets.length === 0) {
        console.log('⚠️  未找到一级资产数据');
        return [];
      }

      console.log(`📊 找到 ${level1Assets.length} 个一级资产，开始查询子资产...`);
      
      const hierarchicalData = [];

      for (let i = 0; i < level1Assets.length; i++) {
        const parentAsset = level1Assets[i];
        console.log(`   ${i + 1}/${level1Assets.length} 处理资产: ${parentAsset['资产编码']} - ${parentAsset['资产名称']}`);

        // 查询该一级资产的子资产
        const childAssets = await this.getChildAssets(parentAsset['资产编码']);
        
        // 构建子资产编码列表（只下一级），每个编码换行
        const childCodes = childAssets.map(child => child['资产编码']).join('\n');
        
        // 添加一级资产记录，包含下级资产编码列表
        const parentRecord = {
          ...parentAsset,
          '下级资产编码列表': childCodes || '',
          '上级资产编码': '' // 一级资产没有上级
        };
        hierarchicalData.push(parentRecord);

        // 添加所有子资产记录，包含上级资产编码
        childAssets.forEach(childAsset => {
          const childRecord = {
            ...childAsset,
            '下级资产编码列表': '', // 子资产没有下级列表
            '上级资产编码': parentAsset['资产编码']
          };
          hierarchicalData.push(childRecord);
        });

        console.log(`     ✅ 找到 ${childAssets.length} 个子资产`);
      }

      console.log(`📈 层级数据构建完成，总计 ${hierarchicalData.length} 条记录`);
      return hierarchicalData;

    } catch (error) {
      console.error('❌ 构建层级数据失败:', error.message);
      throw error;
    }
  }

  /**
   * 生成层级资产Excel文件
   * @returns {string} 生成的文件路径
   */
  async exportToExcel() {
    try {
      const managementAreaName = this.config.managementArea.name;
      const dbName = this.config.databases[this.databaseKey].name;
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
      const filePath = this.excelGenerator.generateHierarchicalExcel(
        hierarchicalData, 
        '十堰东资产层级数据'
      );

      // 显示统计信息
      const level1Count = hierarchicalData.filter(item => item['资产等级'] === 1).length;
      const level2Count = hierarchicalData.filter(item => item['资产等级'] === 2).length;

      console.log('\n📈 导出完成统计:');
      console.log(`   一级资产: ${level1Count} 个`);
      console.log(`   二级资产: ${level2Count} 个`);
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

module.exports = HierarchicalAssetExporter;
