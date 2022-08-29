import pandas as pd
import operator

aa = " "
bb = " "
cc = " "
dd = " "
ee = " "
ff = " "
gg = " "
hh = " "
ini = 0

rdata = pd.read_excel(r"C:\Users\hazyr\OneDrive - dematic.com\Documents\CS OE Email Data - 8.29.2022.xlsx")
columns = ['Path', 'Subj', 'Sender', 'DispTo', 'To', 'Date', 'Body']
rdata.columns = columns

len = operator.length_hint(rdata['Body'])

df = pd.DataFrame({
    'A': [" "],
    'B': [" "],
    'C': [" "],
    'D': [" "],
    'E': [" "],
    'F': [" "],
    'G': [" "],
    'H': [" "],
})

for i in range(len):
    try:
        print(ini)
        x = str(rdata['Body'].iloc[ini])

        a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u = x.split("_x000D_", 20)

        aa,bb,cc,dd,ee,ff,gg,hh = e.split(" / ", 7)



        df.at[ini,'A'] = aa
        df.at[ini,'B'] = bb
        df.at[ini,'C'] = cc
        df.at[ini,'D'] = dd
        df.at[ini,'E'] = ee
        df.at[ini,'F'] = ff
        df.at[ini,'G'] = gg
        df.at[ini,'H'] = hh

        ini = ini + 1


    except:
        ini = ini + 1


df.to_excel("1234567891011.xlsx")

