import numpy as np
import pandas as pd

name="2018190709全市学生成绩表.xlsx"

data = pd.read_excel('D:\\t_01\\tj\\'+name) # 读数据
data1= pd.read_excel(r'D:\t_01\tj\line.xlsx')#提取
data2=pd.read_excel(r'D:\t_01\tj\2_out.xlsx')
data3=pd.read_excel(r'D:\t_01\tj\统计.xlsx')
data7=pd.read_excel(r'D:\t_01\tj\3_out.xlsx')
data8=pd.read_excel(r'D:\t_01\tj\4_out.xlsx')

# 数据清洗与规整
a = list(range(1))
b = list(range(3,4,1))
c=list(range(5,10,2))
d=list(range(13,40,4))
e=list(range(38,39,1))
list_col = a+c+d+e
print (list_col)
df=data.iloc[:,list_col].set_index(['校名',],['校班名称']) #清洗数据结束
e=[]
#更改科目名称
line1=df.columns.values.tolist()
for j in range(len(line1)):
    e.append(line1[j][:2])
ndarray = np.array(e)
print(ndarray)
df.columns=ndarray
#清洗数据查看
df.to_excel(r'D:\t_01\tj_01\result_'+name[:10]+'_d.xlsx')
T_count_col = df.groupby('校班').count()  # 列总计
print(T_count_col)
print(T_count_col.iloc[29:31,:].sum())
print(T_count_col.sum())
Total_count_col = df.groupby('校名').count()  # 列总计
print(Total_count_col)


#Total_count_col.to_excel('D:\\t_01\\tj_01\\'+'line_'+ name)

for i in range(7):
    line=data1.iloc[i].values#循环取分数线
    x=line[11]#取双过线的折总
    x1=line[12]
    print(x)
    tf =df[df.le(line[1:12])]#分数线比对
    tf.to_excel(r'D:\t_01\tj_01\result_18.33s_.xlsx')#比对结果查看
    o_f = (tf[(tf[u'折总'].le(x))])#此处的双过线，必须重新提取line
    #还要提取每个班的选科人数
    oo=o_f.groupby('校班').count()
    oo.insert(0, 'line', x)
    #oo.to_excel('D:\\t_01\\tj_01\\' + str(x) + 'out_file.xlsx')#生成多个单列表
    data2=data2.append(oo,ignore_index=False, sort=False)
oo_data= data2.sort_values(by=['校班', 'line'],axis=0)#双列排序
oo_data.to_excel('D:\\t_01\\tj_01\\'+'oo_'+ name)#生成一个表
y=oo_data._stat_axis.values.tolist()
def find_all_index(arr, item):
    return [i for i, a in enumerate(arr) if a == item]
if __name__ == '__main__':
    a5=[5,4,3,2,1]
    for ii in (data3._stat_axis.values.tolist()):#要统计的校班总数
        a1 = data3.iloc[ii, 0]#循环取出data3中的班号
        a2=(find_all_index(y, a1))#找出在oo_data中的班号位置数组
        print('a2:'+str(a2))
        z=len(a2)#每个班号的位置数组的长度
        #print('z:'+str(z))
        #jj循环出班号的位置

        if z > 5:
            b1 = a5[1:]#4321
            for jj in range(len(b1)):
                #print('jj:' + str(jj))
                g = oo_data.iloc[[a2[jj]]]#* b1[jj]
                if jj==0: g1 = g
                elif jj==1: g2 = g
                elif jj==2: g3 = g
                elif jj==3: g4=g
                else:
                    pass
            gg = g1 * 4 + (g2 - g1) * 3 + (g3 - g2) * 2 + (g4 - g3)
            #data2 = data2.append(oo, ignore_index=False, sort=False)
            data7= data7.append(gg,ignore_index=False, sort=False)
            data7.to_excel(r'D:\t_01\tj_01\101_217'+name)
                #追加完之后，直接写入相应的DataFrame

        else:
            z<=5
            b1 =a5[5-len(a2):]#54321
            for jj in range(len(b1)):
                #print('jj:'+str(jj))
                g = oo_data.iloc[[a2[jj]]]
                if jj == 0:
                    g_1=g
                    g_01 = g_1*b1[jj]
                elif jj ==1:
                    g_2=g
                    g_02 = (g_2-g_1)*b1[jj]
                elif jj == 2:
                    g_3=g
                    g_03 = (g_3-g_2)*b1[jj]
                elif jj == 3:
                    g_4=g
                    g_04= (g_4-g_3)*b1[jj]
                elif jj == 4:
                    g_5=g
                    g_05 = (g_5-g_4)*b1[jj]
                else:
                   pass
            g_11 = g_01 + g_02+ g_03+ g_04 +g_05
            g_01=g_02=g_03= g_04= g_05=0
            data8 = data8.append(g_11,ignore_index=False, sort=False)
            data8.to_excel(r'D:\t_01\tj_01\201_215'+name)





