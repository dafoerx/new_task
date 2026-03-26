# 中医处方管理数据库系统 - 代码实现

## 项目概述
本代码实现基于湖北中医药大学留学生期末考试要求，使用Microsoft Access 2016+开发中医处方管理数据库系统。

## 文件结构
```
code/
├── README.md                    # 本说明文件
├── sql/
│   ├── 01_create_tables.sql    # 创建表结构SQL
│   ├── create_tables.sql       # 创建表结构SQL（备用）
│   ├── create_indexes.sql      # 创建索引SQL
│   ├── create_relationships.sql # 创建关系SQL
│   ├── 03_seed_test_data.sql   # 测试数据插入
│   └── queries/               # 查询实现
│       ├── 1_Count_Physician_Prescriptions_and_Amounts.sql
│       ├── 2_Query_Patients_by_Age_Range.sql
│       ├── 3_Increase_Prescription_Amounts.sql
│       ├── 4_Count_Patients_by_Department_and_Physician.sql
│       └── qryPrescriptionFiveFields.sql # 子窗体查询
├── vba/
│   ├── forms/
│   │   ├── PhysicianBasicInfoForm.cls     # 医师基本信息管理表单
│   │   └── BrowsePrintPrescriptionForm.cls # 浏览与打印处方信息表单
│   ├── reports/
│   │   └── PrintPhysicianPrescriptionReport.cls # 打印医师处方列表报表
│   └── modules/
│       ├── MainModule.bas      # 主程序模块
│       ├── DatabaseUtils.bas   # 数据库工具函数
│       └── UIUtils.bas         # UI工具函数
└── data/
    └── TCM_Prescription_Information.txt  # 处方数据导入文件（示例）
```

## 部署步骤

### 1. 创建Access数据库
1. 打开Microsoft Access 2016或更高版本
2. 创建新的空白数据库，保存为`TCM_Prescription_Management.accdb`
3. 启用宏内容（如果提示安全警告）

### 2. 执行SQL脚本
按顺序执行以下SQL脚本：
1. `sql/create_tables.sql` - 创建三张表
2. `sql/create_indexes.sql` - 创建索引
3. `sql/create_relationships.sql` - 建立表关系
4. `sql/insert_test_data.sql` - 插入测试数据

### 3. 导入VBA代码
1. 按Alt+F11打开VBA编辑器
2. 导入`vba/modules/`目录下的所有.bas文件
3. 导入`vba/forms/`目录下的所有.cls文件（表单类模块）
4. 导入`vba/reports/`目录下的.cls文件（报表类模块）

### 4. 导入外部数据
1. 使用Access的"外部数据"->"文本文件"导入功能
2. 选择`data/TCM_Prescription_Information.txt`文件
3. 导入为新表`TCM_Prescription_Table`
4. 设置字段数据类型：`Physician_ID`为文本，`Patient_Gender`为是/否
5. 确保无重复记录

### 5. 创建查询
在Access中创建以下查询：
1. `1_Count_Physician_Prescriptions_and_Amounts`
2. `2_Query_Patients_by_Age_Range`（参数查询）
3. `3_Increase_Prescription_Amounts`（更新查询）
4. `4_Count_Patients_by_Department_and_Physician`（交叉表查询）

### 6. 创建表单和报表
1. 使用设计视图创建两个表单，应用VBA代码中的UI设置
2. 创建报表，应用格式要求
3. 测试所有功能

## 考试要求验证

### 数据库基础操作（20分）
- [ ] 正确导入`TCM_Prescription_Information.txt`文件
- [ ] 设置`Physician_ID`字段为文本类型
- [ ] 设置`Patient_Gender`字段为是/否类型
- [ ] 确保无重复记录
- [ ] 正确修改表结构
- [ ] 设置`Prescription_ID`为主键
- [ ] 为`Formula_Name`创建普通索引
- [ ] 建立主-从关系，设置级联更新和删除
- [ ] 向`Doctor_Table`插入个人记录

### 查询设计（20分）
- [ ] 查询1：统计姓名恰好三个汉字的医师处方数量和金额
- [ ] 查询2：参数查询，保存结果到新表
- [ ] 查询3：更新查询，将低于150元的处方金额提高5%
- [ ] 查询4：交叉表查询，统计各科室各医师的患者性别分布

### 表单设计（30分）
- [ ] 表单1：医师基本信息管理，对话框样式，无导航控件
- [ ] 表单1：标题标签为隶书22号加粗红色居中
- [ ] 表单1：显示8个字段，细节区域有红色边框
- [ ] 表单1：5个导航按钮功能正常
- [ ] 表单2：浏览与打印处方信息，对话框样式
- [ ] 表单2：标题标签为STXinwei 24号加粗蓝色居中
- [ ] 表单2：包含设计师姓名标签（楷体12号加粗红色）
- [ ] 表单2：子窗体仅显示5个字段
- [ ] 表单2：刷新和预览按钮功能正常

### 综合设计（30分）
- [ ] 报表标题包含医师姓名和职称（宋体24号加粗红色）
- [ ] 报表页脚统计处方数量和总金额（宋体14号加粗红色）
- [ ] 每条处方信息在蓝色边框矩形框内
- [ ] `Prescription_ID`为宋体18号黑色，其他字段宋体12号黑色
- [ ] `Patient_Age`显示为"X岁"，`Prescription_Amount`显示为"¥X元"
- [ ] 处方各区域用点划线分隔（宽度1磅黑色）

## 注意事项
1. **字体要求**：确保系统安装了所需字体（隶书、STXinwei、楷体、宋体）
2. **Access版本**：使用Access 2016或更高版本
3. **提交要求**：将整个项目文件夹重命名为"学号_姓名_考试材料"，压缩后提交
4. **截止时间**：2026年3月27日前提交至指定邮箱

## 快速启动指南

### 一键启动（简化版）
1. 创建新Access数据库 `TCM_Prescription_Management.accdb`
2. 按Alt+F11打开VBA编辑器
3. 导入所有`.bas`和`.cls`文件
4. 执行以下SQL语句（在Access SQL视图中）：

```sql
-- 创建表
CREATE TABLE Department_Table (Department_ID TEXT(10) NOT NULL, Department_Name TEXT(30) NOT NULL);
CREATE TABLE Doctor_Table (Physician_ID TEXT(10) NOT NULL, Department_ID TEXT(10) NOT NULL, Physician_Name TEXT(20) NOT NULL, Physician_Gender TEXT(2) NOT NULL, Physician_Age INTEGER, Physician_Education TEXT(20), Physician_Title TEXT(20));
CREATE TABLE TCM_Prescription_Table (Prescription_ID TEXT(10) NOT NULL, Patient_Name TEXT(10) NOT NULL, Patient_Gender YESNO NOT NULL, Patient_Age INTEGER, Clinical_Diagnosis TEXT(30) NOT NULL, Formula_Name TEXT(30) NOT NULL, Formula_Composition MEMO, Physician_ID TEXT(10) NOT NULL, Prescription_Date DATETIME NOT NULL, Prescription_Amount CURRENCY NOT NULL);

-- 设置主键
ALTER TABLE Department_Table ADD CONSTRAINT PK_Department PRIMARY KEY (Department_ID);
ALTER TABLE Doctor_Table ADD CONSTRAINT PK_Doctor PRIMARY KEY (Physician_ID);
ALTER TABLE TCM_Prescription_Table ADD CONSTRAINT PK_Prescription PRIMARY KEY (Prescription_ID);

-- 插入测试数据
INSERT INTO Department_Table VALUES ('D001', '中医内科');
INSERT INTO Doctor_Table VALUES ('P001', 'D001', '张仲景', '男', 45, '硕士', '主任医师');
INSERT INTO TCM_Prescription_Table VALUES ('RX0001', '李某某', True, 25, '月经不调', '桂枝茯苓丸', '桂枝15g,茯苓20g,芍药15g,丹皮10g,桃仁10g', 'P001', #2015-11-01#, 200);
```

5. 运行 `Main` 过程启动应用程序：
   - 在VBA编辑器中，按F5运行 `Main` 过程
   - 或创建宏调用 `Main` 函数

### 系统测试
运行以下测试验证功能完整性：

1. **数据库测试**：
   ```vba
   ' 在立即窗口中执行
   ? MainModule.CheckDatabaseConnection()
   ```

2. **表单测试**：
   ```vba
   ' 打开医师基本信息管理表单
   DoCmd.OpenForm "frmPhysicianBasicInfo"
   
   ' 打开浏览与打印处方信息表单
   DoCmd.OpenForm "frmBrowsePrintPrescription"
   ```

3. **查询测试**：
   ```vba
   ' 执行四个主要查询
   MainModule.OpenQuery "1_Count_Physician_Prescriptions_and_Amounts"
   ```

## 故障排除

### 常见问题及解决方案

#### 1. 字体不显示或显示异常
- **问题**：表单/报表中的特殊字体（隶书、STXinwei等）显示为默认字体
- **解决方案**：
  - 检查系统是否安装所需字体
  - 在VBA编辑器中运行 `UIUtils.CheckRequiredFonts` 检查字体
  - 如未安装，使用相似字体替代或安装字体包

#### 2. 子窗体不显示数据
- **问题**：浏览与打印处方信息表单的子窗体为空
- **解决方案**：
  - 确保已创建 `qryPrescriptionFiveFields` 查询
  - 检查查询字段名是否与代码一致
  - 确认已选择医师并点击"刷新"按钮

#### 3. 报表格式不符合要求
- **问题**：报表样式与考试要求不一致
- **解决方案**：
  - 运行 `UIUtils.CheckRequiredFonts` 确认字体
  - 检查报表的 `Format` 事件是否正确设置了样式
  - 确认报表控件的名称与VBA代码中引用的名称一致

#### 4. 数据库关系错误
- **问题**：无法建立级联更新/删除关系
- **解决方案**：
  - 确认表已正确创建并包含必要字段
  - 检查外键字段的数据类型是否匹配
  - 在Access关系视图中手动建立关系

#### 5. VBA代码编译错误
- **问题**：导入VBA代码时出现编译错误
- **解决方案**：
  - 确保Access版本为2016或更高
  - 启用VBA项目中的"Microsoft DAO 3.6 Object Library"引用
  - 检查代码中是否有版本不兼容的语法

### 调试技巧

1. **立即窗口调试**：
   ```vba
   ' 在VBA编辑器中按Ctrl+G打开立即窗口
   Debug.Print "变量值：" & variableName
   ```

2. **错误处理调试**：
   - 在所有过程和函数中添加 `On Error GoTo ErrorHandler`
   - 使用 `Err.Number` 和 `Err.Description` 获取错误信息

3. **数据验证**：
   ```vba
   ' 检查数据完整性
   DatabaseUtils.ValidateDataIntegrity
   ```

## 技术支持
如遇到技术问题，请参考：
1. Microsoft Access官方文档
2. VBA编程指南
3. 设计文档中的详细技术说明
4. 代码中的注释和说明

## 提交检查清单
在提交考试作品前，请确认已完成：

- [ ] 数据库文件已重命名为"学号_姓名_TCM_Prescription_Management.accdb"
- [ ] 所有查询按指定名称保存
- [ ] 表单和报表格式符合考试要求
- [ ] 测试数据已正确导入
- [ ] VBA代码无编译错误
- [ ] 文件夹已压缩为.zip或.rar格式
- [ ] 在截止日期前提交至指定邮箱