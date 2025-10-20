#!/usr/bin/env node

const AssetCodeImporter = require('./assetCodeImporter');
const path = require('path');

/**
 * 资产编码对照表导入主程序
 * 使用方法: node importAssetCodes.js [excel文件路径] [是否清空现有数据]
 */
async function main() {
  try {
    // 获取命令行参数
    const args = process.argv.slice(2);
    
    // 解析选项
    let clearExisting = true; // 默认为清空
    let excelFilePath = '新老资产对照.xlsx';
    
    // 处理参数
    for (let i = 0; i < args.length; i++) {
      const arg = args[i];
      if (arg === '--append' || arg === '-a') {
        clearExisting = false;
      } else if (arg === '--clear' || arg === '-c') {
        clearExisting = true;
      } else if (arg === '--help' || arg === '-h') {
        // 帮助信息在下面处理
      } else if (!arg.startsWith('-')) {
        // 如果不是选项，则认为是文件路径
        excelFilePath = arg;
      }
    }

    console.log('='.repeat(60));
    console.log('📊 资产编码对照表导入程序');
    console.log('='.repeat(60));
    console.log(`📁 Excel文件路径: ${excelFilePath}`);
    console.log(`🗑️  清空现有数据: ${clearExisting ? '是' : '否'}`);
    console.log('='.repeat(60));
    
    // 显示使用帮助
    if (args.includes('--help') || args.includes('-h')) {
      console.log('📖 使用说明:');
      console.log('  node importAssetCodes.js [文件路径] [选项]');
      console.log('');
      console.log('选项:');
      console.log('  --append, -a    追加模式（不清空现有数据）');
      console.log('  --clear, -c    清空模式（清空现有数据，默认）');
      console.log('  --help, -h     显示此帮助信息');
      console.log('');
      console.log('示例:');
      console.log('  node importAssetCodes.js                           # 清空模式导入');
      console.log('  node importAssetCodes.js --append                   # 追加模式导入');
      console.log('  node importAssetCodes.js 新老资产对照.xlsx --append # 指定文件，追加模式');
      console.log('='.repeat(60));
      return;
    }

    // 检查文件是否存在
    const fs = require('fs');
    if (!fs.existsSync(excelFilePath)) {
      console.error(`❌ 错误: 文件 "${excelFilePath}" 不存在`);
      console.log('💡 请确保文件路径正确，或使用以下命令:');
      console.log('   node importAssetCodes.js 新老资产对照.xlsx');
      process.exit(1);
    }

    // 创建导入器实例
    const importer = new AssetCodeImporter();

    // 执行导入
    await importer.importFromExcel(excelFilePath, clearExisting);

    console.log('='.repeat(60));
    console.log('✅ 导入完成！');
    console.log('='.repeat(60));

  } catch (error) {
    console.error('❌ 程序执行失败:', error.message);
    console.error('详细错误信息:', error);
    process.exit(1);
  }
}

// 如果直接运行此文件，则执行主程序
if (require.main === module) {
  main();
}

module.exports = { main };
