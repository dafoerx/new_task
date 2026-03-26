-- 查询1：统计医师处方数量与金额
-- 功能：统计姓名恰好为三个汉字的医师的处方数量和金额小计
-- 保存名称：1_Count_Physician_Prescriptions_and_Amounts
-- 分值：5分

SELECT 
    D.Physician_Name AS 医师姓名,
    COUNT(P.Prescription_ID) AS 处方数量,
    SUM(P.Prescription_Amount) AS 处方小计
FROM 
    Doctor_Table AS D
    INNER JOIN TCM_Prescription_Table AS P ON D.Physician_ID = P.Physician_ID
WHERE 
    LEN(D.Physician_Name) = 3  -- 假设姓名为三个汉字
    AND D.Physician_Name NOT LIKE '% %'  -- 确保不包含空格
    AND D.Physician_Name LIKE '[一-龥][一-龥][一-龥]'  -- 确保是三个汉字
GROUP BY 
    D.Physician_Name
ORDER BY 
    处方数量 DESC;

-- 查询说明：
-- 1. 统计每个医师的处方总数和总金额
-- 2. 仅统计姓名恰好为三个汉字的医师
-- 3. 按处方数量降序排列
-- 4. 使用INNER JOIN确保只统计有处方的医师