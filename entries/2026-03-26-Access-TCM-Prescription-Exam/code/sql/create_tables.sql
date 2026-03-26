-- 中医处方管理数据库 - 创建表结构
-- 适用于 Microsoft Access 2016+

-- 1. 科室表 (Department_Table)
CREATE TABLE Department_Table (
    Department_ID TEXT(10) NOT NULL,
    Department_Name TEXT(30) NOT NULL,
    CONSTRAINT PK_Department_Table PRIMARY KEY (Department_ID)
);

-- 2. 医师表 (Doctor_Table)
CREATE TABLE Doctor_Table (
    Physician_ID TEXT(10) NOT NULL,
    Department_ID TEXT(10) NOT NULL,
    Physician_Name TEXT(20) NOT NULL,
    Physician_Gender TEXT(2) NOT NULL,
    Physician_Age INTEGER,
    Physician_Photo IMAGE,
    Physician_Education TEXT(20),
    Physician_Title TEXT(20),
    CONSTRAINT PK_Doctor_Table PRIMARY KEY (Physician_ID)
);

-- 3. 处方表 (TCM_Prescription_Table)
CREATE TABLE TCM_Prescription_Table (
    Prescription_ID TEXT(10) NOT NULL,
    Patient_Name TEXT(10) NOT NULL,
    Patient_Gender YESNO NOT NULL,
    Patient_Age INTEGER,
    Clinical_Diagnosis TEXT(30) NOT NULL,
    Formula_Name TEXT(30) NOT NULL,
    Formula_Composition MEMO,
    Physician_ID TEXT(10) NOT NULL,
    Prescription_Date DATETIME NOT NULL,
    Prescription_Amount CURRENCY NOT NULL,
    CONSTRAINT PK_TCM_Prescription_Table PRIMARY KEY (Prescription_ID)
);

-- 表注释（Access中通过文档属性设置）
-- Department_Table: 存储中医科室信息
-- Doctor_Table: 存储医师基本信息
-- TCM_Prescription_Table: 存储中医处方信息