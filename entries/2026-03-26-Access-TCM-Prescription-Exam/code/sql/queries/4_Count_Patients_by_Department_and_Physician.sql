-- 查询4：统计科室医师患者性别分布
-- 功能：统计各科室各医师治疗的不同性别患者数量，以交叉表形式显示
-- 保存名称：4_Count_Patients_by_Department_and_Physician
-- 分值：5分
-- 类型：交叉表查询

TRANSFORM 
    COUNT(T.Prescription_ID) AS 患者数量
SELECT 
    Dept.Department_Name AS 科室名称,
    D.Physician_Name AS 医师姓名
FROM 
    (Department_Table AS Dept
    INNER JOIN Doctor_Table AS D ON Dept.Department_ID = D.Department_ID)
    INNER JOIN TCM_Prescription_Table AS T ON D.Physician_ID = T.Physician_ID
GROUP BY 
    Dept.Department_Name, D.Physician_Name
PIVOT 
    T.Patient_Gender;

-- 查询说明：
-- 1. 使用TRANSFORM...PIVOT创建交叉表
-- 2. 行标题：科室名称、医师姓名
-- 3. 列标题：患者性别（是/否字段，True/False或0/-1）
-- 4. 值：患者数量（通过COUNT处方ID统计）
-- 5. 使用三层连接：科室→医师→处方

-- 预期输出格式：
-- 科室名称 | 医师姓名 | True（或0） | False（或-1）
-- 内科     | 张三     | 15         | 10
-- 内科     | 李四     | 8          | 12
-- 妇科     | 王五     | 20         | 5

-- 注意事项：
-- 1. Patient_Gender字段是/否类型：0表示男，-1表示女
-- 2. 交叉表会自动将布尔值转换为列标题
-- 3. 可能需要处理空值情况