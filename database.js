const mysql = require('mysql2/promise');

class DatabaseManager {
  constructor(config, databaseKey = 'assets') {
    this.config = config;
    this.databaseKey = databaseKey;
    this.connection = null;
  }

  /**
   * 建立数据库连接
   */
  async connect() {
    try {
      const dbConfig = this.config.databases[this.databaseKey];
      if (!dbConfig) {
        throw new Error(`数据库配置 "${this.databaseKey}" 不存在`);
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
   * 执行查询
   * @param {string} query - SQL查询语句
   * @param {Array} params - 查询参数
   * @returns {Array} 查询结果
   */
  async query(query, params = []) {
    if (!this.connection) {
      await this.connect();
    }

    try {
      const [rows] = await this.connection.execute(query, params);
      console.log(`✅ 查询成功，返回 ${rows.length} 条记录`);
      return rows;
    } catch (error) {
      console.error('❌ 查询失败:', error.message);
      throw error;
    }
  }

  /**
   * 获取资产数据
   * @returns {Array} 资产数据数组
   */
  async getAssetData() {
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
        AS_STATE,
        U_DELETE 
      FROM
        as_asset 
      WHERE
        OPERATING_NAME = '十堰东资产经营中心' 
        AND U_DELETE = 1 
        AND CLASS_TYPE IN ('building', 'structure') 
        AND AS_STATE = 'CHECKED' 
        AND AS_LV = 1 
      ORDER BY
        AS_CODE
    `;

    return await this.query(query);
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

module.exports = DatabaseManager;
