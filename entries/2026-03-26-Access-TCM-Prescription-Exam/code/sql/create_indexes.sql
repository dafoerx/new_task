-- 中医处方管理数据库 - 创建索引
-- 适用于 Microsoft Access 2016+

-- 主键索引（在CREATE TABLE时已定义，这里仅为说明）
-- Department_Table: Department_ID (主键索引)
-- Doctor_Table: Physician_ID (主键索引)
-- TCM_Prescription_Table: Prescription_ID (主键索引)

-- 1. 为TCM_Prescription_Table的Formula_Name字段创建普通索引
-- 提高方剂名称查询性能
CREATE INDEX IDX_Formula_Name ON TCM_Prescription_Table (Formula_Name);

-- 2. 为Doctor_Table的Department_ID字段创建索引（外键性能优化）
CREATE INDEX IDX_Doctor_Department_ID ON Doctor_Table (Department_ID);

-- 3. 为TCM_Prescription_Table的Physician_ID字段创建索引（外键性能优化）
CREATE INDEX IDX_Prescription_Physician_ID ON TCM_Prescription_Table (Physician_ID);

-- 4. 为TCM_Prescription_Table的Prescription_Date字段创建索引（日期查询优化）
CREATE INDEX IDX_Prescription_Date ON TCM_Prescription_Table (Prescription_Date);

-- 索引说明：
-- 1. IDX_Formula_Name: 普通索引，提高按方剂名称查询的性能
-- 2. IDX_Doctor_Department_ID: 外键索引，优化科室-医师关系查询
-- 3. IDX_Prescription_Physician_ID: 外键索引，优化医师-处方关系查询
-- 4. IDX_Prescription_Date: 日期索引，优化按日期范围查询的性能