#loading all sheets
import pandas as pd
sheets=pd.read_excel("Book15.xlsx",sheet_name=None)
print("Menu","1.Search Course","2.Search Certificate","3.Search Course Status",
      "4.Search Course Duration","5.Search Student by Mobile no/admission no",sep="\n")
ch=int(input("Enter your Choice"))
if ch>=1 and ch<=5:
    if ch==1:
        cname=input("enter course name")
        df=pd.DataFrame(sheets["details"])
        list1=df["name"].to_list()
        if cname in list1:
            #p=list1.index(cname)
            print("The given course is available ")
        else:
            print("The given course is not available")
    if ch==2:
        admno=int(input("Enter your admission number"))
        df=pd.DataFrame(sheets["students"])
        list1=df["admnno"].to_list()
        if admno in list1:
            p=list1.index(admno)
            k=df["certificate"].loc[p]
            if k=='y':
                print("certificate already issued")
            else:
                print("certificate not issued")
        else:
            print("admission no invalid")
    if ch==3:
        admno=int(input("Enter your admission number"))
        df=pd.DataFrame(sheets["students"])
        list1=df["admnno"].to_list()
        if admno in list1:
            p=list1.index(admno)
            print(df["status"].loc[p])
        else:
            print("admission number invalid")
    if ch==4:
        cname=input("enter course name")
        df=pd.DataFrame(sheets["details"])
        list1=df["name"].to_list()
        if cname in list1:
            p=list1.index(cname)
            print(df.loc[p])
        else:
            print("course not available")
    if ch==5:
        choi=input("Search student with mobileno/admno")
        if choi== "mobileno":
            mobileno=int(input("enter your mobile no"))
            df=pd.DataFrame(sheets["students"])
            list1=df["mobile"].to_list()
            if mobileno in list1:
                p=list1.index(mobileno)
                print(df["name"].loc[p])
                print(df["status"].loc[p])
            else:
                print("mobile no is not valid")
        if choi=="admno":
            admno=int(input("enter your admn no"))
            df=pd.DataFrame(sheets["students"])
            list1=df["admnno"].to_list()
            if admno in list1:
                p=list1.index(admno)
                print(df["name"].loc[p])
                print(df["status"].loc[p])
            else:
                print("admin no is not valid")
else:
    print("choice is not valid")

