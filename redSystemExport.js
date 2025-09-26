const RedSystemAssetExporter = require('./redSystemAssetExporter');
const config = require('./config');

/**
 * 红色系统资产导出主程序
 */
async function exportRedSystemAssets() {
  // 使用红色系统专用导出器（自动使用红色系统数据库）
  const exporter = new RedSystemAssetExporter(config);
  
  try {
    const managementAreaName = config.managementArea.name;
    const dbName = config.databases.redSystem.name;
    console.log('🚀 开始执行红色系统资产导出任务...');
    console.log(`📋 导出内容：数据库 "${dbName}" 中管理片区 "${managementAreaName}" 的一级资产及其所有子资产的层级关系`);
    
    const filePath = await exporter.exportToExcel();
    
    if (filePath) {
      console.log('\n🎉 红色系统资产导出完成！');
      console.log(`📁 文件位置: ${filePath}`);
    } else {
      console.log('⚠️  没有数据可导出');
    }
    
  } catch (error) {
    console.error('❌ 红色系统资产导出失败:', error.message);
    process.exit(1);
  }
}

// 如果直接运行此文件，则执行导出
if (require.main === module) {
  exportRedSystemAssets();
}

module.exports = { exportRedSystemAssets };
