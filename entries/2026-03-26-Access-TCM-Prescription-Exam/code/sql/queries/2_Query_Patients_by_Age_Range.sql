-- 查询2：按年龄范围查询患者
-- 功能：提示用户输入最小和最大年龄，将25-40岁患者信息保存到新表
-- 保存名称：2_Query_Patients_by_Age_Range
-- 分值：5分
-- 类型：参数查询 + 生成表查询

PARAMETERS 
    [请输入最小年龄:] Long, 
    [请输入最大年龄:] Long;

SELECT 
    Patient_Name AS 患者姓名,
    Patient_Gender AS 性别,
    Patient_Age AS 年龄,
    Clinical_Diagnosis AS 临床诊断
INTO 
    [Patients Aged 25 to 40]
FROM 
    TCM_Prescription_Table
WHERE 
    Patient_Age BETWEEN [请输入最小年龄:] AND [请输入最大年龄:];

-- 使用示例：
-- 1. 运行此查询时，Access会提示"请输入最小年龄:"和"请输入最大年龄:"
-- 2. 输入25和40（按题目要求）
-- 3. 查询将创建新表"Patients Aged 25 to 40"
-- 4. 新表包含患者姓名、性别、年龄、临床诊断四个字段

-- 注意事项：
-- 1. 参数名称必须与题目要求完全一致
-- 2. 新表名称必须为"Patients Aged 25 to 40"（包含空格）
-- 3. 确保年龄字段为数字类型