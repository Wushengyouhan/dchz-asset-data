const HierarchicalAssetExporter = require('./hierarchicalAssetExporter');
const config = require('./config');

/**
 * 蓝色系统层级资产导出主程序
 */
async function exportHierarchicalAssets() {
  // 使用红色系统数据库查询蓝色系统数据
  const exporter = new HierarchicalAssetExporter(config, 'redSystem');
  
  try {
    const managementAreaName = config.managementArea.name;
    console.log('🚀 开始执行层级资产导出任务...');
    console.log(`📋 导出内容：管理片区 "${managementAreaName}" 的一级资产及其所有子资产的层级关系`);
    
    const filePath = await exporter.exportToExcel();
    
    if (filePath) {
      console.log('\n🎉 层级资产导出完成！');
      console.log(`📁 文件位置: ${filePath}`);
    } else {
      console.log('⚠️  没有数据可导出');
    }
    
  } catch (error) {
    console.error('❌ 层级资产导出失败:', error.message);
    process.exit(1);
  }
}

// 如果直接运行此文件，则执行导出
if (require.main === module) {
  exportHierarchicalAssets();
}

module.exports = { exportHierarchicalAssets };
