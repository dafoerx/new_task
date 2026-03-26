-- Access SQL: 测试数据
-- 如需重复导入，可先清空子表后再清空父表。

INSERT INTO Department_Table (Department_ID, Department_Name)
VALUES ('D001', '中医内科');

INSERT INTO Department_Table (Department_ID, Department_Name)
VALUES ('D002', '中医妇科');

INSERT INTO Department_Table (Department_ID, Department_Name)
VALUES ('D003', '针灸科');

INSERT INTO Department_Table (Department_ID, Department_Name)
VALUES ('D004', '推拿科');

INSERT INTO Doctor_Table
    (Physician_ID, Department_ID, Physician_Name, Physician_Gender, Physician_Age, Physician_Education, Physician_Title)
VALUES
    ('P001', 'D001', '张仲景', '男', 45, '硕士', '主任医师');

INSERT INTO Doctor_Table
    (Physician_ID, Department_ID, Physician_Name, Physician_Gender, Physician_Age, Physician_Education, Physician_Title)
VALUES
    ('P002', 'D002', '李时珍', '男', 52, '博士', '主任医师');

INSERT INTO Doctor_Table
    (Physician_ID, Department_ID, Physician_Name, Physician_Gender, Physician_Age, Physician_Education, Physician_Title)
VALUES
    ('P003', 'D002', '王玉华', '女', 38, '硕士', '副主任医师');

INSERT INTO Doctor_Table
    (Physician_ID, Department_ID, Physician_Name, Physician_Gender, Physician_Age, Physician_Education, Physician_Title)
VALUES
    ('P004', 'D003', '赵明轩', '男', 41, '本科', '主治医师');

INSERT INTO Doctor_Table
    (Physician_ID, Department_ID, Physician_Name, Physician_Gender, Physician_Age, Physician_Education, Physician_Title)
VALUES
    ('P005', 'D004', '陈静怡', '女', 36, '硕士', '主治医师');

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000001', '李某某', True, 25, '月经不调', '桂枝茯苓丸', '桂枝15g,茯苓20g,芍药15g,丹皮10g,桃仁10g', 'P003', #2015-11-01#, 200);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000002', '王小明', False, 34, '胃脘痛', '香砂六君子汤', '木香6g,砂仁6g,党参15g,白术12g,茯苓15g,甘草6g', 'P001', #2015-11-02#, 148);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000003', '赵婷婷', True, 29, '失眠', '酸枣仁汤', '酸枣仁20g,茯苓15g,知母10g,川芎6g,甘草6g', 'P001', #2015-11-03#, 132);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000004', '孙建国', False, 47, '颈肩疼痛', '针灸处方A', '百会,风池,肩井,曲池,合谷', 'P004', #2015-11-03#, 120);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000005', '周丽娜', True, 31, '产后调理', '八珍汤', '党参12g,白术12g,茯苓15g,甘草6g,当归10g,川芎6g,白芍12g,熟地12g', 'P002', #2015-11-04#, 168);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000006', '马天宇', False, 40, '腰肌劳损', '推拿处方B', '腰背推拿20分钟,拔罐10分钟,艾灸15分钟', 'P005', #2015-11-05#, 98);

INSERT INTO TCM_Prescription_Table
    (Prescription_ID, Patient_Name, Patient_Gender, Patient_Age, Clinical_Diagnosis, Formula_Name, Formula_Composition, Physician_ID, Prescription_Date, Prescription_Amount)
VALUES
    ('RX00000007', '何春华', True, 43, '更年期综合征', '逍遥散', '柴胡10g,白芍12g,当归10g,白术12g,茯苓15g,炙甘草6g,薄荷3g', 'P003', #2015-11-06#, 186);
