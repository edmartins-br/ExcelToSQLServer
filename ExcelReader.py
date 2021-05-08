import pandas as pd
import pyodbc

print('Criando conexão com banco...')
conn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=XX.XXX.XXX.XX;'
                      'DATABASE=XXXXXX;'
                      'UID=XXXXX;'
                      'PWD=XXXXX;'
                      'Trusted_Connection=no;')
cursor = conn.cursor()

print('Lendo arquivo Excel...')
data = pd.read_excel(r'.\excel\ativos.xlsx')
df = pd.DataFrame(data, columns = ['Month', 'Year', 'Local_ID',
                                   'Global_ID','Employee_name',
                                   'Gender','Birthday','Age','fs',
                                   'Hiring','Seniority','Staff_Operator',
                                   'Grade','Subgrade','Job_Family',
                                   'Career_level','Brazil_job_code','cbo',	
                                   'Base_salary','Trust_position', 	
                                   'Regular_Compatible','pcd','Status',
                                   'Status_type','Email_address',
                                   'Employment_type','Subsidiary',	
                                   'org_Product','Accounting_unit',	
                                   'Expense_type','org_Portuguese',	
                                   'org_English','org_Code',	
                                   'org_Name_GERP','org_Code',	
                                   'org_Name_RHEV','Location',	
                                   'Manager','Assistant','Realocated',	
                                   'Cont_Index','Index'])
df = df.fillna("")
print(df)

print('Criando TABELA...')
cursor.execute('''CREATE TABLE headCount ([Month] INT, [Year] INT, LocalID INT, Global_ID VARCHAR(50), 
                EmployeeName VARCHAR(200), Gender VARCHAR(50), Birthday datetimeoffset(7),
                Age int, fs VARCHAR(50),
                Hiring datetimeoffset(7), Seniority int, Staff_Operator VARCHAR(50),
                Grade VARCHAR(50), Subgrade VARCHAR(50), JobFamily VARCHAR(50),
                CareerLevel VARCHAR(50), BrazilJobCode VARCHAR(50), cbo int,
                BaseSalary DECIMAL(11,2), TrustPosition VARCHAR(50), Regular_Compatible VARCHAR(50),
                pcd VARCHAR(50), Status VARCHAR(50), StatusType VARCHAR(50), EmailAddres VARCHAR(200),
                EmploymentType VARCHAR(50), Subsidiary NVARCHAR(50), OrgProduct VARCHAR(50),
                AccounttingUnit VARCHAR(50), ExpenseType VARCHAR(50), OrgPortuguese VARCHAR(50),
                OrgEnglish VARCHAR(50), OrgCode int, ORGName VARCHAR(50), 
                ORGCode VARCHAR(50), ORGName NVARCHAR(50), [Location] NVARCHAR(50), Manager NVARCHAR(50), Assistant NVARCHAR(50),
                Realocated VARCHAR(50), ContIndex VARCHAR(50), [Index] VARCHAR(50))''')

print('Inserindo dados no banco banco...')

#insert dataFrame into the Table
for row in df.itertuples():    
    cursor.execute('''
                INSERT INTO HeadCountStaging.dbo.headCount ([Month] , [Year] , LocalID , [Global_ID], EmployeeName, Gender,
                                                            Birthday, [Age], fse_ise_k_ise, Hiring, [Seniority], Staff_Operator, Grade, Subgrade,
                                                            JobFamily, CareerLevel, BrazilJobCode, [cbo], BaseSalary, TrustPosition,
                                                            Regular_Compatible, pcd, Status, StatusType, EmailAddres, EmploymentType,
                                                            Subsidiary, OrgProduct, AccounttingUnit, ExpenseType, OrgPortuguese,
                                                            OrgEnglish, [OrgCodeGerp], ORGNameGERP, ORGCodeRHEV, ORGNameRHEV,  Location, Manager, Assistant,
                                                            Realocated, ContIndex, [Index])
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''',
                row.Month,
                row.Year, 
                row.Local_ID, 
                row.Global_ID, 
                row.Employee_name, 
                row.Gender,
                row.Birthday, 
                row.Age, 
                row.fse_ise_k_ise, 
                row.Hiring, 
                row.Seniority, 
                row.Staff_Operator, 
                row.Grade, 
                row.Subgrade,
                row.Job_Family, 
                row.Career_level, 
                row.Brazil_job_code, 
                row.cbo, 
                row.Base_salary, 
                row.Trust_position,
                row.Regular_Compatible, 
                row.pcd, 
                row.Status, 
                row.Status_type, 
                row.Email_address, 
                row.Employment_type,
                row.Subsidiary, 
                row.org_Product, 
                row.Accounting_unit, 
                row.Expense_type, 
                row.org_Portuguese,
                row.org_English, 
                row.org_Code_GERP, 
                row.org_Name_GERP, 
                row.org_Code_RHEV,
                row.org_Name_RHEV,
                row.Location, 
                row.Manager, 
                row.Assistant,
                row.Realocated, 
                row.Cont_Index, 
                row.Index
                )
print('DADOS INSERIDOS COM SUCESSO!')
conn.commit()

