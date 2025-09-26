// 多数据库配置文件示例
// 请复制此文件为 config.js 并填入您的实际配置

module.exports = {
  // 数据库配置 - 支持多个数据库
  databases: {
    // 蓝色系统
    assets: {
      name: '蓝色系统',
      host: 'mr-pamfuqcb968wxc7xyo-vpc.rwlb.rds.aliyuncs.com',
      port: 5999,
      user: 'hdbs_assets',
      password: 'hd3bsE3tcdeNsAst2s',
      database: 'hdbs_assets',
      charset: 'utf8mb4',
      timezone: '+08:00',
      dateStrings: true,
      sqlInjectionDefense: true
    },
    // 红色系统
    redSystem: {
      name: '红色系统',
      host: 'mr-pamfuqcb968wxc7xyo-vpc.rwlb.rds.aliyuncs.com',
      port: 5999,
      user: 'hdbs_dchz_user',
      password: 'ShGxNXaCqmMRjR3f',
      database: 'hdbs_dchz',
      charset: 'utf8mb4',
      timezone: '+08:00',
      dateStrings: true,
      sqlInjectionDefense: true,
      login_encryption_algorithm: 'sha256'
    }
  },
  
  // 当前使用的数据库
  currentDatabase: 'assets', // 可选值: 'assets', 'redSystem'
  
  excel: {
    outputDir: './output'
  },
  
  // 管理片区配置
  managementArea: {
    name: '十堰东资产经营中心'  // 修改为其他管理片区名称
  },
  
  // 资产对照表配置
  assetComparison: {
    // 蓝色系统文件名（可选，如果不指定则自动查找最新文件）
    blueSystemFile: null, // 例如: '蓝色系统_十堰东资产经营中心_资产数据_2025-09-25T01-30-11.xlsx'
    // 红色系统文件名（可选，如果不指定则自动查找最新文件）
    redSystemFile: null,  // 例如: '红色系统_十堰东资产经营中心_资产数据_2025-09-25T02-11-05.xlsx'
    // 是否自动查找最新文件（当指定文件名为null时）
    autoFindLatest: true
  }
};
