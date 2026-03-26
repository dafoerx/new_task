-- 中医处方管理数据库 - 创建表关系
-- 适用于 Microsoft Access 2016+
-- 注意：在Access中，关系通常通过数据库工具->关系视图创建
-- 以下SQL语句在Access中可能不完全支持，作为指导使用

-- 1. Department_Table 与 Doctor_Table 的一对多关系
-- 关联字段：Department_Table.Department_ID → Doctor_Table.Department_ID
-- 参照完整性：强制
ALTER TABLE Doctor_Table
ADD CONSTRAINT FK_Doctor_Department
FOREIGN KEY (Department_ID)
REFERENCES Department_Table (Department_ID);

-- 2. Doctor_Table 与 TCM_Prescription_Table 的一对多关系
-- 关联字段：Doctor_Table.Physician_ID → TCM_Prescription_Table.Physician_ID
-- 参照完整性：级联更新相关字段、级联删除相关字段
ALTER TABLE TCM_Prescription_Table
ADD CONSTRAINT FK_Prescription_Physician
FOREIGN KEY (Physician_ID)
REFERENCES Doctor_Table (Physician_ID)
ON UPDATE CASCADE
ON DELETE CASCADE;

-- 关系说明：
-- 1. 一个科室可以有多个医师，一个医师属于一个科室
-- 2. 一个医师可以开具多个处方，一个处方由一个医师开具
-- 3. 级联更新：当医师ID更新时，相关处方自动更新
-- 4. 级联删除：当医师记录删除时，相关处方自动删除