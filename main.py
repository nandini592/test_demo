import os
import pandas as pd

os.chdir("C:\\bytexl\\SIDDHANT")
files = os.listdir()

# ##Finding K for College Name
df = pd.DataFrame()
def drop(f):
    df = pd.read_excel(f,sheet_name=1,index_col=0)
    #index_no = df.columns.get_loc('Q1')
    df.columns = df.iloc[0]
    df = df[1:]
    df.drop(df.columns[10:], axis = 1, inplace = True)
    df['College Name'] = 'College Name'
    return df


# drop_loc = df.loc['Q1']
# print(drop_loc)

def time_convert(df):
    t = df['Time'].tolist()
    for i in range(len(t)):
        a = t[i].split()
        if(len(a)==2):
            if(int(a[0][:-1])>=15):
                t[i] = 15
            else:
                c,d = a[0][:-1],a[1][:-1]
                c = int(c)
                d = int(d)
                t[i] = round(c+d/60,2)
        elif(len(a)==1) and (a[0] == '-'):
            t[i]=0
        elif(len(a)==1):
            c = a[0][:-1]
            c = int(c)
            t[i] = round(c/60,2)
        else:
            t[i]=15
    df["Time"] = t 
    df.index.names = ['User ID']
    return df

def topic_exam_tech_non_tech(df,f):
    l = f.split("-")
    #df["Year"] = l[-1].split(".")[0]
    df["Year"] = '2023'
    df["Topic"] = l[1]
    df["Exam"]  = l[2].split(".")[0]
    
    if((l[1].lower() == "aptitude") or (l[1].lower() == "verbal") or (l[1].lower() == "verabal")):
        df["Tech or Non-Tech"] = "Non-Tech"
    else:
        df["Tech or Non-Tech"] = "Tech"
    return df

def fileMerge():
    
    df = pd.DataFrame()
    df1 = pd.DataFrame()
    for f in files:
        df1 = pd.read_excel(f)
        df = df.append(df1,ignore_index=True)
    return df

#Driver Code Below
def dataMerge(df1,df2):
    merge = pd.merge(df1, df2, on = "Email",how = "left")
    return merge
if __name__ == "__main__":
    os.chdir("C:\\bytexl\\SIDDHANT")
    files = os.listdir()
    global collegeName
    collegeName = (files[0].split("-")[0])
    for file in files: 
        df = pd.DataFrame() 
        df = time_convert(drop(file))
        topic_exam_tech_non_tech(df,file)
    df.head()
    merged_df = fileMerge()
    os.chdir("C:\\bytexl")#directory in which student data is located in
    student_df = pd.read_excel('KJCE STUDENT DATA.xlsx')#student data
    os.chdir("C:\\bytexl\\test\\testing")
    final_df = dataMerge(merged_df,student_df)
    #final_df.drop_duplicates(subset=['Email','User ID','Date','Score','Exam','Topic','Time'])
    os.chdir("C:\\bytexl\\test\\testing")#target directory
    final_df.to_excel(collegeName+".xlsx")
    