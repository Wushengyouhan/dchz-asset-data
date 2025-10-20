const ContractDataProcessor = require('./contractDataProcessor');
const path = require('path');
const fs = require('fs');

/**
 * 主函数：处理合同管理片区数据
 */
async function main() {
  try {
    console.log('🚀 开始处理合同管理片区数据...');
    console.log('📋 目标：十堰西资产经营中心（仅处理十堰西的老资产）');
    
    // 创建处理器实例
    const processor = new ContractDataProcessor();
    
    // 输入文件路径（十堰西资产编码文件）
    const inputFilePath = './十堰西资产编码.xlsx';
    
    // 检查输入文件是否存在
    if (!fs.existsSync(inputFilePath)) {
      console.error(`❌ 输入文件不存在: ${inputFilePath}`);
      console.log('💡 请确保Excel文件路径正确');
      return;
    }
    
    // 生成输出文件路径
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputDir = './output';
    
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    const outputFilePath = path.join(outputDir, `十堰西资产经营中心_新资产信息_${timestamp}.xlsx`);
    
    console.log(`📖 输入文件: ${inputFilePath}`);
    console.log(`📊 输出文件: ${outputFilePath}`);
    
    // 处理Excel文件
    await processor.processExcelFile(inputFilePath, outputFilePath);
    
    console.log('\n🎉 处理完成！');
    console.log(`📁 结果文件已保存至: ${outputFilePath}`);
    
  } catch (error) {
    console.error('❌ 程序执行失败:', error.message);
    console.error('详细错误信息:', error);
    process.exit(1);
  }
}

// 如果直接运行此文件，则执行主函数
if (require.main === module) {
  main().catch(error => {
    console.error('❌ 程序执行失败:', error);
    process.exit(1);
  });
}

module.exports = { main };
