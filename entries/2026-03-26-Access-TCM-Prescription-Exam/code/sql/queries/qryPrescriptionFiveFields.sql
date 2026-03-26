-- 查询：处方五字段查询
-- 用途：为浏览与打印处方信息表单的子窗体提供数据源
-- 仅显示5个字段：患者姓名、临床诊断、方剂名称、处方日期、处方金额
-- 同时包含医师ID用于筛选

SELECT 
    P.Prescription_ID,
    P.Patient_Name,
    P.Clinical_Diagnosis,
    P.Formula_Name,
    P.Prescription_Date,
    P.Prescription_Amount,
    P.Physician_ID
FROM 
    TCM_Prescription_Table AS P
ORDER BY 
    P.Prescription_Date DESC;

-- 查询说明：
-- 1. 包含Prescription_ID用于内部标识（不在子窗体中显示）
-- 2. 包含Physician_ID用于与主窗体关联（LinkChildFields）
-- 3. 显示5个字段：Patient_Name, Clinical_Diagnosis, Formula_Name, Prescription_Date, Prescription_Amount
-- 4. 按处方日期降序排列，最新的处方显示在最前面

-- 在Access中创建此查询的步骤：
-- 1. 打开查询设计视图
-- 2. 添加TCM_Prescription_Table表
-- 3. 选择上述7个字段
-- 4. 设置Prescription_Date为降序排序
-- 5. 保存查询名称为"qryPrescriptionFiveFields"

-- 注意事项：
-- 1. 确保TCM_Prescription_Table表已存在并包含数据
-- 2. 查询名称必须与VBA代码中的名称一致
-- 3. 如果字段名有变化，需要相应修改此查询和VBA代码