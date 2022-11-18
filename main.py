'''
Instructions:
1. click the green arrow in the top right corner next to main
2. after program is done running open the file it will be on the left and will be named the date and a random number
3. clear out the OE Sheets in the attachments folder
------------------------------------------------------------------------------------------------------------------------------------
If there are any issues, my name is Ryan Hazy, I was a rotational finance intern at the time I made this. There is a good chance
that I do not work here anymore. So give me a call at (616) 550-0684.
'''

# This just imports the libraries that are used during this program
from datetime import datetime
import numpy as np
import pandas as pd
import os
import random

# ini moves from one OE sheet to the next
ini=0
# wi is for the output file. It keeps track of what line things are getting written to
wi = 0

# This goes through the specified folder where the saved attachments reside
entries = os.listdir(r"C:\Users\hazyr\OneDrive - dematic.com\Attachments")

# this is setting up the dataframe for the output
#these are the colloumns
oucols = ['b', 'project number', 'customer', 'b', 'b', 'profit center', 'b', 'salesman', 'b', 'new business', 'margin', 'b', 'category']
output = pd.DataFrame(columns=oucols)

#each time this loop runs a new OE sheet is being processed
for i in range(len(entries)):
    # curfil is the current file name and that comes from the entries variable above
    curfil = entries[ini]
    # this is the path of the folder where the files reside
    filpat = r"C:\Users\hazyr\OneDrive - dematic.com\Attachments"
    # this combines the file path and name to create a complete address so the file can be used by the program
    filnam=os.path.join(filpat,curfil)

    # this reads the file and grabs the OE Sheet tab
    oefile = pd.read_excel(filnam, sheet_name= "OE Sheet")
    oefile = pd.DataFrame(oefile)
    # these what the coloumns are being named
    collumns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB']
    oefile.columns = collumns

    # this is grabbing the sheet with the first project name
    otheroe = pd.read_excel(filnam, sheet_name=19)
    otheroe = pd.DataFrame(otheroe)
    colomns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ']
    otheroe.columns = colomns

    # this goes through the OE Sheet line by line
    ix = 0

    # this is grabbing the data from the sheet and the numbers are displaced by 2
    # project number
    pronum = oefile.iloc[ix+34]['B']
    # customer name
    cust = oefile.iloc[3]['F']
    # profit center
    procen = oefile.iloc[ix+34]['N']
    # sales person
    salper = oefile.iloc[10]['B']
    # new business
    newbus = oefile.iloc[ix+34]['C']
    # category
    category = oefile.iloc[ix+34]['I']
    # margin
    margin = ((otheroe.iloc[1]['B']) - (otheroe.iloc[0]['B']))

    # this if statement decides whether or nor it should be written in the out put
    if (newbus > 0 or newbus < 0):
        # this is writing the previous variable in the output statement
        output.at[wi,'project number'] = pronum
        output.at[wi,'customer'] = cust
        output.at[wi,'profit center'] = procen
        output.at[wi,'salesman'] = salper
        output.at[wi,'new business'] = newbus
        output.at[wi,'margin'] = margin
        output.at[wi,'category'] = category
        # this is the line that the output gets written on increasing
        wi = wi + 1
    # si shows how many tabs over from 19 the sheet with the project name is on
    si = 1
    # this is the row counter for the OE Sheet increasing by one
    ix = ix + 1

    # this goes through the sheet and looks for the stuff.
    # The part above this could probably be combined with this part, but the program works and there is no point in changing now
    for i in range(40):
        pronum = oefile.iloc[ix + 34]['B']
        cust = oefile.iloc[3]['F']
        procen = oefile.iloc[ix + 34]['N']
        salper = oefile.iloc[10]['B']
        newbus = oefile.iloc[ix + 34]['C']
        category = oefile.iloc[ix + 34]['I']

        # this is looking to make sure the correct project file is being looked at
        try:
            otheroe = pd.read_excel(filnam, sheet_name=19+si)
            otheroe = pd.DataFrame(otheroe)
            colomns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R','S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI','AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX','AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ']
            otheroe.columns = colomns
            margin = ((otheroe.iloc[1]['B']) - (otheroe.iloc[0]['B']))
        except:
            otheroe = pd.read_excel(filnam, sheet_name=19)
            otheroe = pd.DataFrame(otheroe)
            colomns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S','T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ','AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ','BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ']
            otheroe.columns = colomns
            margin = ((otheroe.iloc[1]['B']) - (otheroe.iloc[0]['B']))

            # this is writing to the output file
        if (newbus > 0 or newbus < 0):
            output.at[wi, 'project number'] = pronum
            output.at[wi, 'customer'] = cust
            output.at[wi, 'profit center'] = procen
            output.at[wi, 'salesman'] = salper
            output.at[wi, 'new business'] = newbus
            output.at[wi, 'margin'] = margin
            output.at[wi, 'category'] = category
            wi = wi + 1

        ix = ix + 1

    ini = ini + 1

# this is getting the current date to be used for the file name
now = datetime.now()
date = str(now.date())
# this is a random number generator to make sure that files do not get named the same
rn = "-" + str(random.randint(100,999))
fname = date + rn + ".xlsx"

#this section is writing the output to excel
writer = pd.ExcelWriter(fname)
output.to_excel(writer)
writer.save()
writer.close()

# this is so you can see what the random number attached to your file is
print(rn)
