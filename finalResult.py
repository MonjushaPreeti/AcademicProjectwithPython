import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

col = [0, 1, 2, 3, 4, 5, 6, 7];
c_101 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=59, usecols=col, nrows=62);
# print(c_101);
temp1 = [];
# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_101.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 12:
        k = i;
        a = int(abs(c_101.iloc[k:k + 1]['Internal'] - c_101.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_101.iloc[k:k + 1]['External'] - c_101.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp1.append(int(c_101.iloc[k:k + 1]['Internal'] + c_101.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp1.append(int(c_101.iloc[k:k + 1]['External'] + c_101.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp1.append(int(c_101.iloc[k:k + 1]['Internal'] + c_101.iloc[k:k + 1]['External']) / 2);
c_101_final = [];
c_101['Total'] = c_101['Average'] + c_101['Tutorial(40)'];
for i in c_101['Total']:
    if i >= 80:
        c_101_final.append(4.00);
    elif i >= 75:
        c_101_final.append(3.75);
    elif i >= 70:
        c_101_final.append(3.50);
    elif i >= 65:
        c_101_final.append(3.25);
    elif i >= 60:
        c_101_final.append(3.00);
    elif i >= 55:
        c_101_final.append(2.75);
    elif i >= 50:
        c_101_final.append(2.50);
    elif i >= 45:
        c_101_final.append(2.25);
    elif i >= 40:
        c_101_final.append(2.00);
    else:
        c_101_final.append(1.00);
print(c_101_final);
col = [0, 1, 2, 3, 4, 5, 6, 7];
c_103 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=125, usecols=col, nrows=62);
# print(c_101);
temp2 = [];
# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_103.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 12:
        k = i;
        a = int(abs(c_103.iloc[k:k + 1]['Internal'] - c_103.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_103.iloc[k:k + 1]['External'] - c_103.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp2.append(int(c_103.iloc[k:k + 1]['Internal'] + c_103.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp2.append(int(c_103.iloc[k:k + 1]['External'] + c_103.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp2.append(int(c_103.iloc[k:k + 1]['Internal'] + c_103.iloc[k:k + 1]['External']) / 2);
c_103_final = [];
c_103['Total'] = c_103['Average'] + c_103['Tutorial'];
for i in c_103['Total']:
    if i >= 80:
        c_103_final.append(4.00);
    elif i >= 75:
        c_103_final.append(3.75);
    elif i >= 70:
        c_103_final.append(3.50);
    elif i >= 65:
        c_103_final.append(3.25);
    elif i >= 60:
        c_103_final.append(3.00);
    elif i >= 55:
        c_103_final.append(2.75);
    elif i >= 50:
        c_103_final.append(2.50);
    elif i >= 45:
        c_103_final.append(2.25);
    elif i >= 40:
        c_103_final.append(2.00);
    else:
        c_103_final.append(1.00);
print(c_103_final);
col = [0, 1, 2, 3, 4, 5, 6, 7];
c_105 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=192, usecols=col, nrows=62);
# print(c_101);
temp3 = [];

for i in range(62):
    j = int(c_101.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 12:
        k = i;
        a = int(abs(c_105.iloc[k:k + 1]['Internal'] - c_105.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_105.iloc[k:k + 1]['External'] - c_105.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp3.append(int(c_105.iloc[k:k + 1]['Internal'] + c_105.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp3.append(int(c_105.iloc[k:k + 1]['External'] + c_105.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp3.append(int(c_105.iloc[k:k + 1]['Internal'] + c_105.iloc[k:k + 1]['External']) / 2);
c_105_final = [];
c_105['Total'] = c_105['Average'] + c_105['Tutorial'];
for i in c_105['Total']:
    if i >= 80:
        c_105_final.append(4.00);
    elif i >= 75:
        c_105_final.append(3.75);
    elif i >= 70:
        c_105_final.append(3.50);
    elif i >= 65:
        c_105_final.append(3.25);
    elif i >= 60:
        c_105_final.append(3.00);
    elif i >= 55:
        c_105_final.append(2.75);
    elif i >= 50:
        c_105_final.append(2.50);
    elif i >= 45:
        c_105_final.append(2.25);
    elif i >= 40:
        c_105_final.append(2.00);
    else:
        c_105_final.append(1.00);
print(c_105_final);
col = [0, 1, 2, 3, 4, 5, 6, 7];
c_107 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=259, usecols=col, nrows=62);
# print(c_101);
temp4 = [];
# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_107.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 12:
        k = i;
        a = int(abs(c_107.iloc[k:k + 1]['Internal'] - c_107.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_107.iloc[k:k + 1]['External'] - c_107.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp4.append(int(c_107.iloc[k:k + 1]['Internal'] + c_107.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp4.append(int(c_107.iloc[k:k + 1]['External'] + c_107.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp4.append(int(c_107.iloc[k:k + 1]['Internal'] + c_107.iloc[k:k + 1]['External']) / 2);
c_107_final = [];
c_107['Total'] = c_107['Average'] + c_107['Tutorial'];
for i in c_107['Total']:
    if i >= 80:
        c_107_final.append(4.00);
    elif i >= 75:
        c_107_final.append(3.75);
    elif i >= 70:
        c_107_final.append(3.50);
    elif i >= 65:
        c_107_final.append(3.25);
    elif i >= 60:
        c_107_final.append(3.00);
    elif i >= 55:
        c_107_final.append(2.75);
    elif i >= 50:
        c_107_final.append(2.50);
    elif i >= 45:
        c_107_final.append(2.25);
    elif i >= 40:
        c_107_final.append(2.00);
    else:
        c_107_final.append(1.00);
print(c_107_final);
col = [0, 1, 2, 3, 4, 5, 6, 7];
c_109 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=326, usecols=col, nrows=62);
# print(c_101).;
temp5 = [];
# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_109.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 12:
        k = i;
        a = int(abs(c_109.iloc[k:k + 1]['Internal'] - c_109.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_109.iloc[k:k + 1]['External'] - c_109.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp5.append(int(c_109.iloc[k:k + 1]['Internal'] + c_109.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp5.append(int(c_109.iloc[k:k + 1]['External'] + c_109.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp5.append(int(c_109.iloc[k:k + 1]['Internal'] + c_109.iloc[k:k + 1]['External']) / 2);
c_109_final = [];
c_109['Total'] = c_109['Average'] + c_109['Tutorial'];
for i in c_109['Total']:
    if i >= 80:
        c_109_final.append(4.00);
    elif i >= 75:
        c_109_final.append(3.75);
    elif i >= 70:
        c_109_final.append(3.50);
    elif i >= 65:
        c_109_final.append(3.25);
    elif i >= 60:
        c_109_final.append(3.00);
    elif i >= 55:
        c_109_final.append(2.75);
    elif i >= 50:
        c_109_final.append(2.50);
    elif i >= 45:
        c_109_final.append(2.25);
    elif i >= 40:
        c_109_final.append(2.00);
    else:
        c_109_final.append(1.00);
print(c_109_final);
col = [0, 1, 2, 3, 4, 5, 6, 7];
c_111 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=393, usecols=col, nrows=62);
# print(c_101).;
temp6 = [];
# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_111.iloc[i:i + 1]['Difference']);
    k = i;
    if j > 6:
        k = i;
        a = int(abs(c_111.iloc[k:k + 1]['Internal'] - c_111.iloc[k:k + 1]['3rd Examiner']));
        b = int(abs(c_111.iloc[k:k + 1]['External'] - c_111.iloc[k:k + 1]['3rd Examiner']));
        if a < b:
            temp6.append(int(c_111.iloc[k:k + 1]['Internal'] + c_111.iloc[k:k + 1]['3rd Examiner']) / 2);
        else:
            temp6.append(int(c_111.iloc[k:k + 1]['External'] + c_111.iloc[k:k + 1]['3rd Examiner']) / 2);
    else:
        temp6.append(int(c_111.iloc[k:k + 1]['Internal'] + c_111.iloc[k:k + 1]['External']) / 2);
c_111_final = [];
c_111['Total'] = c_111['Average'] + c_111['Tutorial'];
for i in c_111['Total']:
    if i * 2 >= 80:
        c_111_final.append(4.00);
    elif i * 2 >= 75:
        c_111_final.append(3.75);
    elif i * 2 >= 70:
        c_111_final.append(3.50);
    elif i * 2 >= 65:
        c_111_final.append(3.25);
    elif i * 2 >= 60:
        c_111_final.append(3.00);
    elif i * 2 >= 55:
        c_111_final.append(2.75);
    elif i * 2 >= 50:
        c_111_final.append(2.50);
    elif i * 2 >= 45:
        c_111_final.append(2.25);
    elif i * 2 >= 40:
        c_111_final.append(2.00);
    else:
        c_111_final.append(1.00);
print(c_111_final);
col1 = [0, 1, 2, 3, 4];
c_108 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=461, usecols=col1, nrows=62);
# print(c_101).;

# temp2 = [];
# temp3 = [];
for i in range(62):
    j = int(c_108.iloc[i:i + 1]['CSE-108'])
c_108_final = [];
for i in c_108['CSE-108']:
    if i >= 80:
        c_108_final.append(4.00);
    elif i >= 75:
        c_108_final.append(3.75);
    elif i >= 70:
        c_108_final.append(3.50);
    elif i >= 65:
        c_108_final.append(3.25);
    elif i >= 60:
        c_108_final.append(3.00);
    elif i >= 55:
        c_108_final.append(2.75);
    elif i >= 50:
        c_108_final.append(2.50);
    elif i >= 45:
        c_108_final.append(2.25);
    elif i >= 40:
        c_108_final.append(2.00);
    else:
        c_108_final.append(1.00);
print(c_108_final);
col2 = [0, 1, 2, 3, 4];
c_114 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=529, usecols=col2, nrows=63);
# print(c_101).;

# temp2 = [];
# temp3 = [];
temp7 = [];
for i in range(62):
    temp7.append(float(c_108.iloc[i:i + 1]['Class test'] + c_108.iloc[i:i + 1]['Final']));
c_114_final = [];
for i in c_114['CSE-114']:
    if i * 2 >= 80:
        c_114_final.append(4.00);
    elif i * 2 >= 75:
        c_114_final.append(3.75);
    elif i * 2 >= 70:
        c_114_final.append(3.50);
    elif i * 2 >= 65:
        c_114_final.append(3.25);
    elif i * 2 >= 60:
        c_114_final.append(3.00);
    elif i * 2 >= 55:
        c_114_final.append(2.75);
    elif i * 2 >= 50:
        c_114_final.append(2.50);
    elif i * 2 >= 45:
        c_114_final.append(2.25);
    elif i * 2 >= 40:
        c_114_final.append(2.00);
    else:
        c_114_final.append(1.00);
print(c_114_final);
col3 = [0, 1, 2];
c_100 = pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls", skiprows=598, usecols=col2, nrows=62);
# print(c_101).;

# temp2 = [];
# temp3 = [];
c_100_final = [];
for i in c_100['CSE-100']:
    if i * 2 >= 80:
        c_100_final.append(4.00);
    elif i * 2 >= 75:
        c_100_final.append(3.75);
    elif i * 2 >= 70:
        c_100_final.append(3.50);
    elif i * 2 >= 65:
        c_100_final.append(3.25);
    elif i * 2 >= 60:
        c_100_final.append(3.00);
    elif i * 2 >= 55:
        c_100_final.append(2.75);
    elif i * 2 >= 50:
        c_100_final.append(2.50);
    elif i * 2 >= 45:
        c_100_final.append(2.25);
    elif i * 2 >= 40:
        c_100_final.append(2.00);
    else:
        c_100_final.append(1.00);
print(c_100_final);
# df=pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls");
# df['Total_credits']=df['c__100_final']*1+df['c_101_final']*3+df['c_103_final']*3+df['c_105_final']*3+df['c_107_final']*3+df['c_108_final']*1+df['c_109_final']*3+df['c_111_final']*2+df['c_114_final']*1;
# df['CGPA']=round(df['Total_credits']/20,2);
# print(df['CGPA'].describe);
total_cgpa = [];
for i in range(62):
    cgpa = (c_100_final[i] * 1 + c_101_final[i] * 3 + c_103_final[i] * 3 + c_105_final[i] * 3 + c_107_final[i] * 3 +
            c_108_final[i] * 1 + c_109_final[i] * 3 + c_111_final[i] * 2 + c_114_final[i] * 1);
    # total_cgpa.append()
    total_cgpa.append(round(cgpa / 20, 2));
print(total_cgpa)
# df=pd.read_excel("C:\\Users\\user\\Downloads\\Year-1-Sem-1-2016.xls");
df = pd.DataFrame({'CGPA': total_cgpa},index=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,11,12,13,14,15
                  ,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,
                  42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61])
df1 = pd.DataFrame({'Roll':[ 150028,150060,160001,160002,160003,160004,160005,160006,160007,160008,160009,160010,160011,160012,160013,160014,160015,160016,160017
,160018,160019,160020,160021,160022,160023,160024,160025,160026,160027,160028,160029,160030,160031,160032,160033,160034,160035,160036,160037,160038,160039
,160040,160041,160042,160043,160044,160045,160046,160047,160048,160049,160050,160051,160052,160053,160054,160055,160056,160057,160058,160059
,160060]})

writer = pd.ExcelWriter("C:\\Users\\user\\Downloads\\R_result.xls");
df.to_excel(writer, 'Sheet1', index=False,startrow=0,startcol=4);
df1.to_excel(writer, 'Sheet1',index=False,startrow=0,startcol=3);
writer.save();
p=int(input("Enter the roll number"));
print(total_cgpa[p-159999]);
