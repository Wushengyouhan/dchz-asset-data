const DatabaseManager = require('./database');
const ExcelGenerator = require('./excelGenerator');
const config = require('./config');

/**
 * 主函数：从数据库读取数据并生成Excel文件
 */
async function main() {
  const dbManager = new DatabaseManager(config.database);
  const excelGenerator = new ExcelGenerator(config);

  try {
    console.log('🚀 开始执行数据导出任务...');
    console.log('📋 查询条件：十堰东资产经营中心 - 建筑物和构筑物 - 已审核 - 一级资产');

    // 1. 连接数据库
    await dbManager.connect();

    // 2. 查询资产数据
    console.log('🔍 正在查询数据库...');
    const assetData = await dbManager.getAssetData();

    if (assetData.length === 0) {
      console.log('⚠️  未找到符合条件的数据');
      return;
    }

    // 3. 生成Excel文件
    console.log('📊 正在生成Excel文件...');
    const filePath = excelGenerator.generateExcel(assetData, '十堰东资产数据');

    // 4. 显示结果统计
    console.log('\n📈 导出完成统计:');
    console.log(`   总记录数: ${assetData.length}`);
    console.log(`   文件路径: ${filePath}`);
    console.log(`   文件大小: ${require('fs').statSync(filePath).size} bytes`);

    // 5. 显示前几条数据预览
    console.log('\n👀 数据预览 (前3条):');
    assetData.slice(0, 3).forEach((item, index) => {
      console.log(`   ${index + 1}. ${item['资产编码']} - ${item['资产名称']} (${item['资产分类']})`);
    });

  } catch (error) {
    console.error('❌ 执行过程中发生错误:', error.message);
    process.exit(1);
  } finally {
    // 关闭数据库连接
    await dbManager.close();
  }
}

// 如果直接运行此文件，则执行主函数
if (require.main === module) {
  main().catch(error => {
    console.error('❌ 程序执行失败:', error);
    process.exit(1);
  });
}

module.exports = { main, config };
