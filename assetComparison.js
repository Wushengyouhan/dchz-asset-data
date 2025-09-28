const AssetComparisonGenerator = require('./assetComparisonGenerator');
const config = require('./config');
const fs = require('fs');
const path = require('path');

/**
 * 资产对照表生成主程序
 */
async function generateAssetComparison() {
  const generator = new AssetComparisonGenerator(config);
  
  try {
    console.log('🚀 开始生成资产对照表...');
    
    // 查找最新的蓝色系统和红色系统Excel文件
    const outputDir = config.excel.outputDir;
    const managementAreaName = config.managementArea.name;
    
    console.log('🔍 正在查找Excel文件...');
    
    let blueFile, redFile;
    
    // 检查是否指定了具体的文件名
    if (config.assetComparison.blueSystemFile && config.assetComparison.redSystemFile) {
      // 使用配置的文件名
      blueFile = {
        name: config.assetComparison.blueSystemFile,
        path: path.join(outputDir, config.assetComparison.blueSystemFile)
      };
      redFile = {
        name: config.assetComparison.redSystemFile,
        path: path.join(outputDir, config.assetComparison.redSystemFile)
      };
      
      console.log(`📁 使用配置的蓝色系统文件: ${blueFile.name}`);
      console.log(`📁 使用配置的红色系统文件: ${redFile.name}`);
      
      // 检查文件是否存在
      if (!fs.existsSync(blueFile.path)) {
        throw new Error(`蓝色系统文件不存在: ${blueFile.path}`);
      }
      if (!fs.existsSync(redFile.path)) {
        throw new Error(`红色系统文件不存在: ${redFile.path}`);
      }
    } else {
      // 自动查找最新文件
      console.log('🔍 自动查找最新的Excel文件...');
      
      // 查找蓝色系统文件（资产管理系统）
      const blueSystemFiles = fs.readdirSync(outputDir)
        .filter(file => file.startsWith('蓝色系统_') && file.includes(managementAreaName) && file.includes('资产数据_'))
        .map(file => ({
          name: file,
          path: path.join(outputDir, file),
          time: fs.statSync(path.join(outputDir, file)).mtime
        }))
        .sort((a, b) => b.time - a.time);
      
      // 查找红色系统文件
      const redSystemFiles = fs.readdirSync(outputDir)
        .filter(file => file.startsWith('红色系统_') && file.includes(managementAreaName) && file.includes('资产数据_'))
        .map(file => ({
          name: file,
          path: path.join(outputDir, file),
          time: fs.statSync(path.join(outputDir, file)).mtime
        }))
        .sort((a, b) => b.time - a.time);
      
      if (blueSystemFiles.length === 0) {
        throw new Error(`未找到蓝色系统文件，请先运行: npm run hierarchical`);
      }
      
      if (redSystemFiles.length === 0) {
        throw new Error(`未找到红色系统文件，请先运行: npm run red-system`);
      }
      
      blueFile = blueSystemFiles[0];
      redFile = redSystemFiles[0];
      
      console.log(`📁 自动选择蓝色系统文件: ${blueFile.name}`);
      console.log(`📁 自动选择红色系统文件: ${redFile.name}`);
    }
    
    // 读取Excel文件
    console.log('📖 正在读取Excel文件...');
    const blueAssets = generator.readExcelFile(blueFile.path);
    const redAssets = generator.readExcelFile(redFile.path);
    
    // 生成两个对照表数据
    console.log('📊 正在生成对照表数据...');
    const blueSystemComparisonData = generator.generateBlueSystemComparisonData(blueAssets, redAssets);
    const redSystemComparisonData = generator.generateRedSystemComparisonData(blueAssets, redAssets);
    
    // 生成蓝色系统对照表Excel文件
    console.log('📊 正在生成蓝色系统对照表Excel文件...');
    const blueSystemFilePath = generator.generateComparisonExcel(blueSystemComparisonData, '蓝色系统对照表');
    
    // 生成红色系统对照表Excel文件
    console.log('📊 正在生成红色系统对照表Excel文件...');
    const redSystemFilePath = generator.generateComparisonExcel(redSystemComparisonData, '红色系统对照表');
    
    // 显示统计信息
    const blueMatchedCount = blueSystemComparisonData.filter(item => item['匹配状态'] === '已匹配').length;
    const blueUnmatchedCount = blueSystemComparisonData.filter(item => item['匹配状态'] === '未匹配').length;
    const redMatchedCount = redSystemComparisonData.filter(item => item['匹配状态'] === '已匹配').length;
    const redUnmatchedCount = redSystemComparisonData.filter(item => item['匹配状态'] === '未匹配').length;
    
    console.log('\n🎉 双对照表生成完成！');
    console.log(`📁 蓝色系统对照表: ${blueSystemFilePath}`);
    console.log(`📁 红色系统对照表: ${redSystemFilePath}`);
    console.log(`\n📊 统计信息:`);
    console.log(`   蓝色系统对照表:`);
    console.log(`     总计: ${blueSystemComparisonData.length} 条`);
    console.log(`     已匹配: ${blueMatchedCount} 条`);
    console.log(`     未匹配: ${blueUnmatchedCount} 条`);
    console.log(`   红色系统对照表:`);
    console.log(`     总计: ${redSystemComparisonData.length} 条`);
    console.log(`     已匹配: ${redMatchedCount} 条`);
    console.log(`     未匹配: ${redUnmatchedCount} 条`);
    
    return { blueSystemFilePath, redSystemFilePath };
    
  } catch (error) {
    console.error('❌ 资产对照表生成失败:', error.message);
    process.exit(1);
  }
}

// 如果直接运行此文件，则执行生成
if (require.main === module) {
  generateAssetComparison();
}

module.exports = { generateAssetComparison };
