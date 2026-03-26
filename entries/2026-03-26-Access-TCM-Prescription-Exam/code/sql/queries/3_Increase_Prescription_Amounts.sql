-- 查询3：调整处方金额
-- 功能：将TCM_Prescription_Table中金额低于150元的处方金额提高5%
-- 保存名称：3_Increase_Prescription_Amounts
-- 分值：5分
-- 类型：更新查询

UPDATE 
    TCM_Prescription_Table
SET 
    Prescription_Amount = Prescription_Amount * 1.05
WHERE 
    Prescription_Amount < 150;

-- 查询说明：
-- 1. 仅更新金额低于150元的处方
-- 2. 将金额提高5%（乘以1.05）
-- 3. 金额等于或高于150元的处方保持不变

-- 执行前数据示例：
-- 处方A：金额100元 → 执行后：105元
-- 处方B：金额149元 → 执行后：156.45元
-- 处方C：金额150元 → 执行后：150元（不变）
-- 处方D：金额200元 → 执行后：200元（不变）

-- 注意事项：
-- 1. 使用<而不是<=，因为题目要求"低于150元"
-- 2. 货币类型计算会自动保留两位小数
-- 3. 建议在执行前备份数据