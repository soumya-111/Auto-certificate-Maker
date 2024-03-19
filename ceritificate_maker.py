import pandas as pd
import os
import time
import tqdm as tk
import simple_colors as clr
import colorama
# from fontTools import font
# from textblob import TextBlob
import markdown
from colorama import Fore,Back,Style
colorama.init()
from PIL import Image,ImageFont,ImageDraw
def excel_alldata(x,y,sr,er,cr,ff,fontdz):
    df=pd.read_excel(x)
    clmn_data=list(df.columns)
    clmn_data=str(clmn_data[0])
    df[clmn_data] = df[clmn_data].astype("str")
    df1=pd.read_excel("d:\\anime.xlsx")
    excellist1=df.values.tolist()
    excellist2=df1.values.tolist()

    if (sr == None and er == None and cr == None):
        data = [i for i in excellist1]
    if (sr != None and er != None and cr == None):
        data = [i for i in excellist1 if i[0] >= x1 and i[0] <= y1]
    if (sr != None and er == None and cr == None):
        data = [i for i in excellist1 if sr == i[0]]
    if (sr == None and er == None and cr != None):
        data = [i for i in excellist1 if i[0] in cr]

    if os.path.exists(ff):
        pass
    else:
        os.mkdir(ff)

    for i in range(len(data)):
        img = Image.open(y)
        I1 = ImageDraw.Draw(img)
        IMW,IMH=img.size
        for j in range(len(excellist2)):
            X=0
            Y=0
            myfont = ImageFont.truetype(f'C:\\Users\\MY PC\\AppData\\Local\\Microsoft\\Windows\\Fonts\\{fontdz}.ttf', excellist2[j][4])
            #C:\\Windows\\Fonts
            # myfont = ImageFont.truetype(f'C:\\Windows\\Fonts\\{fontdz}.ttf',excellist2[j][4])
            # \\Windows\\Fonts
            xldata=str(data[i][j])
            dx,dy= I1.textsize(xldata,myfont)
            if(excellist2[j][3]==0):
                X=excellist2[j][1]
                Y=excellist2[j][2]
            if (excellist2[j][3] == 1):
                if(excellist2[j][1] == 0):
                    X = (IMW - dx) / 2
                    Y=excellist2[j][2]
                if(excellist2[j][2] == 0):
                    X=excellist2[j][1]
                    Y = (IMH - dy) / 2
            if(excellist2[j][3]==2):
                X=(IMW-(IMW-excellist2[j][1])-dx)
                Y = excellist2[j][2]
                # X = TextBlob(X)
                # X.set_bold(True)
                # Y = TextBlob(Y)
                # Y.set_bold(True)
                # font = font.Font("GreatVibes-Regular.ttf")
                # font.setBold(True)
                # X=font.getGlyph(Y)
                # Y=font.getGlyph(Y)
                X=markdown.markdown(X)
                Y=markdown.markdown(Y)
            I1.text((X ,Y),xldata, font=myfont, fill="#"+excellist2[j][5])

        # f=str(data[i][0])
        # ff1=ff+"\\"+f
        # if os.path.exists(ff1):
        #     pass
        # else:
        #     os.mkdir(ff1)

        # gmname= int=len(data[0])-1
        img = img.convert('RGB')
        # image.save('image.jpg')
        img.save(ff+"\\"  + data[i][0]+".pdf")

def design():
    print(clr.blue("Processing............."))
    time.sleep(1.0)
    for i in tk.tqdm(range(100), desc=clr.blue("File generating......: ")):
        time.sleep(.1)
    print()
    print(clr.blue("Successfully Completed.............\n"))
    print()

a1="\t\t\t\t\t\t\t\t........Welcome to Ajantrik World........\n\n"
for i in a1:
    print('\033[1m' + clr.red(i),end="")
    if(i==" "):
        time.sleep(.7)
    time.sleep(.0199)
print("\033[0;0m")


print(clr.red("Press 1 :---------------> Certificate for all Student."))
print(clr.yellow("Press 2 :---------------> Certificate for given range Students."))
print(clr.cyan("Press 3 :---------------> Certificate for single Student."))
print(clr.green("Press 4 :---------------> Certificate for random reg no. Student."))
print(clr.blue("Press 5 :---------------> EXIT\n\n"))

while(True):
    ch=int(input(clr.green("Enter your choice :")))
    if(ch==1):
        x="d:\\FInalanime.xlsx"
        y="d:\\AKIMATSURI Certificate.jpg"
        ff="d:\\animenew109"
        # x=input(clr.red("Enter excel(data) path :"))
        # y=input(clr.yellow("Enter certificate template path :"))
        # ff=input(clr.blue("Enter Folder create path :"))
        # fontdz=input(clr.magenta("Enter font style name :"))
        fontdz="GreatVibes-Regular"
        #GreatVibes - Regular
        excel_alldata(x,y,None,None,None,ff,fontdz)
        print()
        design()
    elif(ch==2):
        x = "d:\\Data for template (3).xlsx"
        y = "d:\\Entry Card.png"
        x1="AKI0022"
        y1="AKI0032"
        ff = "d:\\animenew3"
        fontdz = "ARIALBD"
        # x1 = input(clr.red("Enter lower record (reg no.) :"))
        # y1 = input(clr.blue("Enter upper record (reg no.) :"))
        # x = input(clr.yellow("Enter excel(data) path :"))
        # y = input(clr.green("Enter certificate template path :"))
        # ff = input(clr.magenta("Enter Folder create path :"))
        # fontdz = input(clr.cyan("Enter font style name :"))

        excel_alldata(x,y,x1,y1,None,ff,fontdz)
        print()
        design()
    elif(ch==3):
        x = "d:\\Data for template (11).xlsx"
        y = "d:\\Entry Card.png"
        ff = "d:\\animesingle1"
        x2="SRISTI SHAW"
        fontdz = "ARIALBD"
        # x2 = input(clr.red("Enter single student reg no:"))
        # x = input(clr.blue("Enter excel(data) path :"))
        # y = input(clr.cyan("Enter certificate template path :"))
        # ff = input(clr.magenta("Enter Folder create path :"))
        # fontdz = input(clr.yellow("Enter font style name :"))
        excel_alldata(x,y,x2,None,None,ff,fontdz)
        print()
        design()
    elif(ch==5):
        print(clr.blue("Exit Successfully."))
        exit()
    elif(ch==4):
        # x = "d:\\a.xlsx"
        # y = "d:\\R.png"
        # ff = "d:\\new13"
        # s1="1000001","2400010"
        s1 = input(clr.red("Enter student reg no.(use coma) :")).split(",")
        x = input(clr.green("Enter excel(data) path :"))
        y = input(clr.blue("Enter certificate template path :"))
        ff = input(clr.magenta("Enter Folder create path :"))
        fontdz = input(clr.yellow("Enter font style name :"))
        s2=[i for i in s1]
        excel_alldata(x,y,None,None,s2,ff,fontdz)
        print()
        design()
    else:
        print()
        print(clr.cyan("Enter valid input.............\n"))
