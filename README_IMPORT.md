# 资产编码对照表导入程序

## 功能说明

这个程序用于将Excel文件中的老资产编码和新资产编码对照关系导入到MySQL数据库中。

## 文件结构

- `assetCodeImporter.js` - 核心导入器类
- `importAssetCodes.js` - 主程序入口
- `新老资产对照.xlsx` - 数据源Excel文件

## 数据库表结构

程序会自动创建以下数据库表：

```sql
CREATE TABLE `old_as_code_new` (
  `ID` varchar(32) NOT NULL COMMENT '用户ID',
  `OLD_AS_CODE` varchar(20) NOT NULL COMMENT '老资产编码',
  `NEW_AS_CODE` varchar(80) NOT NULL COMMENT '新资产编码',
  PRIMARY KEY (`ID`),
  UNIQUE KEY `unique_old_as_code` (`OLD_AS_CODE`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=DYNAMIC COMMENT='老资产对照表';
```

## 使用方法

### 1. 基本使用

```bash
# 使用默认文件名导入
node importAssetCodes.js

# 指定Excel文件路径
node importAssetCodes.js 新老资产对照.xlsx

# 不清空现有数据（追加模式）
node importAssetCodes.js 新老资产对照.xlsx false
```

### 2. 使用npm脚本

```bash
# 使用npm脚本运行
npm run import-codes
```

## Excel文件格式要求

Excel文件必须包含以下列：
- `OLD_AS_CODE` - 老资产编码
- `NEW_AS_CODE` - 新资产编码

示例数据：
```
OLD_AS_CODE                    NEW_AS_CODE
DT-SY-W-1050643-100-100       DT-SY-W-1050643-01
DT-SY-W-1050644-100-100       DT-SY-W-1050644-01
```

## 程序特性

1. **自动建表**: 如果数据库表不存在，程序会自动创建
2. **数据验证**: 导入前会验证Excel文件格式和数据完整性
3. **重复处理**: 如果老资产编码已存在，会更新对应的新资产编码
4. **错误处理**: 完善的错误处理和日志记录
5. **数据验证**: 导入完成后会验证数据完整性

## 配置说明

程序使用 `config.js` 中的数据库配置，确保数据库连接信息正确。

## 注意事项

1. 确保Excel文件格式正确，包含必要的列
2. 确保数据库连接配置正确
3. 程序默认会清空现有数据，如需保留现有数据请使用追加模式
4. 老资产编码必须唯一，重复的编码会被更新

## 错误处理

程序包含完善的错误处理机制：
- 文件不存在检查
- 数据库连接检查
- 数据格式验证
- 导入过程监控

## 日志输出

程序会输出详细的执行日志，包括：
- 文件读取状态
- 数据解析结果
- 数据库操作状态
- 导入结果验证
