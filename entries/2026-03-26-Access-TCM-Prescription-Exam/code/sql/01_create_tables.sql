-- Access SQL: 创建核心表结构、键关系与索引
-- 建议在 Access 的“创建 > 查询设计 > SQL 视图”中逐段执行。

CREATE TABLE Department_Table (
    Department_ID TEXT(10) NOT NULL,
    Department_Name TEXT(30) NOT NULL,
    CONSTRAINT PK_Department_Table PRIMARY KEY (Department_ID)
);

CREATE TABLE Doctor_Table (
    Physician_ID TEXT(10) NOT NULL,
    Department_ID TEXT(10) NOT NULL,
    Physician_Name TEXT(20) NOT NULL,
    Physician_Gender TEXT(2) NOT NULL,
    Physician_Age INTEGER,
    Physician_Photo OLEOBJECT,
    Physician_Education TEXT(20),
    Physician_Title TEXT(20),
    CONSTRAINT PK_Doctor_Table PRIMARY KEY (Physician_ID),
    CONSTRAINT FK_Doctor_Department
        FOREIGN KEY (Department_ID)
        REFERENCES Department_Table (Department_ID)
);

CREATE TABLE TCM_Prescription_Table (
    Prescription_ID TEXT(10) NOT NULL,
    Patient_Name TEXT(10) NOT NULL,
    Patient_Gender YESNO NOT NULL,
    Patient_Age INTEGER,
    Clinical_Diagnosis TEXT(30) NOT NULL,
    Formula_Name TEXT(30) NOT NULL,
    Formula_Composition LONGTEXT,
    Physician_ID TEXT(10) NOT NULL,
    Prescription_Date DATETIME NOT NULL,
    Prescription_Amount CURRENCY NOT NULL,
    CONSTRAINT PK_TCM_Prescription_Table PRIMARY KEY (Prescription_ID),
    CONSTRAINT FK_Prescription_Doctor
        FOREIGN KEY (Physician_ID)
        REFERENCES Doctor_Table (Physician_ID)
        ON UPDATE CASCADE
        ON DELETE CASCADE
);

-- 索引：设计文档要求 Formula_Name 普通索引
CREATE INDEX IX_TCM_Prescription_Formula_Name
    ON TCM_Prescription_Table (Formula_Name);

-- 推荐索引：加速常见关联查询
CREATE INDEX IX_Doctor_Department_ID
    ON Doctor_Table (Department_ID);

CREATE INDEX IX_TCM_Prescription_Physician_ID
    ON TCM_Prescription_Table (Physician_ID);
