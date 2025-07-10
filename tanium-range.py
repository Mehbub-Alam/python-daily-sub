import pandas as pd
from tkinter import *
from tkinter import filedialog
import time
from datetime import datetime, timedelta


ts = time.time()

# Add date range - if you are running this today, add today's date
dateChosenOlder = "5/1/2025"
dateChosenNewer = "5/2/2025"

dateNoneOlder = datetime.strptime(dateChosenOlder, "%m/%d/%Y")
dateNoneNewer = datetime.strptime(dateChosenNewer, "%m/%d/%Y")

myDateOlder = dateNoneOlder - timedelta(days=7)
myDateNewer = dateNoneNewer - timedelta(days=7)

# Date minus 7 days ago
try:
    finalDateOlder = myDateOlder.strftime("%-m/%-d/%Y") #Linux OS
    finalDateNewer = myDateNewer.strftime("%-m/%-d/%Y") #Linux OS
except ValueError:
    finalDateOlder = myDateOlder.strftime("%#m/%#d/%Y") # Windows OS
    finalDateNewer = myDateNewer.strftime("%#m/%#d/%Y") # Windows OS

print("finalDateOlder is of \n")
print(type(finalDateOlder))
print(finalDateOlder + "\n")
# print(finalDateNewer)

print("finalDateNewer is of \n")
print(type(finalDateNewer))
print(finalDateNewer + "\n")

# GUI input taker
# GUI input taker
root = Tk()
root.withdraw()
clientStatusFile = filedialog.askopenfilename(title="Select Client Status file", filetypes=(("csv files", ".csv"), ("all files", ".*")))
rawReportsFile = filedialog.askopenfilename(title="Select Raw Report file", filetypes=(("csv files", ".csv"), ("all files", ".*")))
cmdbStatusFile = filedialog.askopenfilename(title="Select CMDB Status file", filetypes=(("csv files", ".csv"), ("all files", ".*")))
awsStatusFile = filedialog.askopenfilename(title="Select AWS Status file", filetypes=(("csv files", ".csv"), ("all files", ".*")))

# root.destroy()

root.wm_deiconify()
label = Label(root, text="Running the operation", font=('Arial', 12, 'bold'), bg=f'#e1edf5', fg='#ff003c', width=50, height=2)
label.pack()
root.mainloop()
pd.options.display.max_rows = 9999999


# Read CSV files into Dataframe
readInput = pd.read_csv(clientStatusFile, usecols=["Host Name","Last Registration"])
readInput2 = pd.read_csv(rawReportsFile, usecols=["Computer Name","Last Seen", "Operating System"])
readInput3 = pd.read_csv(cmdbStatusFile, encoding='windows-1252', usecols=["name","install_status", "operational_status"])
# readInput4 = pd.read_csv(awsStatusFile, encoding='windows-1252', usecols=["Name","State"])
readInput4 = pd.read_csv(awsStatusFile, usecols=["Name","State"])

df = pd.DataFrame(readInput)
df2= pd.DataFrame(readInput2)
df3= pd.DataFrame(readInput3)
df4= pd.DataFrame(readInput4)

print(readInput4.head())

# Loop to read raw server report and sort by date
df2["Date2"] = ""
for index, cell in df2.iterrows():
    df2.at[index, "Computer Name"] = cell["Computer Name"].split(".")[0]
    # if "AIX" in cell["Operating System"]:
    #     # print(cell["Operating System"])
    #     df2.drop(index, axis=0, inplace=True)
    df2.at[index, "Date2"] = regDate = pd.to_datetime(cell["Last Seen"].split("T")[0])

# Drop duplicates
df2.sort_values(by="Date2", inplace=True, ascending=False)
df2.drop_duplicates(subset="Computer Name", keep="first", inplace=True)

# new 0
# print("Checking Anomaly raw server report: ")
# print(df2)

#### Loop to read client status report and sort drop duplicates ####################
df["Date"] = ""
for index, cell in df.iterrows():
    df.at[index, "Host Name"] = cell["Host Name"].split(".")[0]
    df.at[index, "Date"] = pd.to_datetime(cell["Last Registration"].split(",")[0])


df.sort_values(by="Date", inplace=True, ascending=False)
df.drop_duplicates(subset=['Host Name'], keep="first", inplace=True)

# new
# print("Checking Anomaly Client Status: ")
# print(df)

#check Date range and drop other rows //* WE CAN"T COMPARE STRING DATE * Change it to Date Object - FIXED
for index, cell in df.iterrows():
    comparedDate = cell["Last Registration"].split(",")[0]
    comparedDateConverted = datetime.strptime(comparedDate, "%m/%d/%Y")
    finalDateOlderConverted = datetime.strptime(finalDateOlder, "%m/%d/%Y")
    finalDateNewerConverted = datetime.strptime(finalDateNewer, "%m/%d/%Y")
    # if ((cell["Last Registration"].split(",")[0] < finalDateOlder) | (cell["Last Registration"].split(",")[0] > finalDateNewer)):
    if ((comparedDateConverted < finalDateOlderConverted) | (comparedDateConverted > finalDateNewerConverted)):
        df.drop(index, axis=0, inplace=True)

# new 2
# print("Checking Anomaly Client status after date stripped: ")
# print(df)

# Merge the dataframes for
dfTemp = df.merge(df2, left_on=df["Host Name"].str.lower(), right_on=df2["Computer Name"].str.lower(), how="inner")[["Computer Name", "Last Registration", "Operating System","Last Seen"]]

dfTemp2 = dfTemp.merge(df3, how="left",left_on=dfTemp["Computer Name"].str.lower(), right_on=df3["name"].str.lower())[["Computer Name", "Last Registration", "Operating System","Last Seen", "install_status", "operational_status"]]


dfOutput = dfTemp2.merge(df4, how="left", left_on=dfTemp2["Computer Name"].str.lower(), right_on=df4["Name"].str.lower())[["Computer Name", "Last Registration", "Operating System","Last Seen", "install_status", "operational_status","State"]]

dfOutput.drop_duplicates(subset="Computer Name", keep="first", inplace=True)
dfOutput = dfOutput.fillna("N/A")

# print("Checking Anomaly: ")
# print(dfOutput)

exportPath = clientStatusFile.rsplit("/", 1)

dfOutput.to_excel(f'{exportPath[0]}/Cleaned combinedRaw_{int(ts)}.xlsx', index=False)
print("Operation is successful")


for index, cell in dfOutput.iterrows():
    if ((cell["operational_status"] == "N/A") and (cell["install_status"] == "N/A") and (cell["State"] == "N/A")):
    #     # print(cell["Operating System"])
        dfOutput.drop(index, axis=0, inplace=True)

linuxCount = 0
windowsCount = 0

#Adding OS Type
dfOutput["OS Type"] = ""
for index, cell in dfOutput.iterrows():
    if ("AIX" in cell["Operating System"]) or ("Linux" in cell["Operating System"]):
        dfOutput.at[index, "OS Type"] = "Linux"
        linuxCount+=1
    else:
        dfOutput.at[index, "OS Type"] = "Windows"
        windowsCount+=1

# Get counts

cmdbWindowsInactiveCount=0
cmdbLinuxInactiveCount=0
awsWindowsInactiveCount = 0
awsLinuxInactiveCount = 0
windowsTicketCount = 0
linuxTicketCount = 0



for index, cell in dfOutput.iterrows():
    if ((cell["operational_status"] != "Operational") and (cell["OS Type"] == "Windows")):
        cmdbWindowsInactiveCount+=1
    if ((cell["operational_status"] != "Operational") and (cell["OS Type"] == "Linux")):
        cmdbLinuxInactiveCount+=1
    if ((cell["State"] == "stopped") and (cell["OS Type"] == "Windows")):
        awsWindowsInactiveCount+=1
    if ((cell["State"] == "stopped") and (cell["OS Type"] == "Linux")):
        awsLinuxInactiveCount+=1

    if ((cell["operational_status"] == "Operational") and (cell["State"] != "stopped") and (cell["OS Type"] == "Windows")):
        windowsTicketCount+=1
    if ((cell["operational_status"] == "Operational") and (cell["State"] != "stopped") and (cell["OS Type"] == "Linux")):
        linuxTicketCount+=1

   


print("Windows Count: "+str(windowsCount))
print("Linux Count: "+str(linuxCount))
print("CMDB Windows Inactive Count: "+str(cmdbWindowsInactiveCount))
print("CMDB Linux Inactive Count: "+str(cmdbLinuxInactiveCount))
print("AWS Windows Inactive Count: "+str(awsWindowsInactiveCount))
print("AWS Linux Inactive Count: "+str(awsLinuxInactiveCount))
print("Windows Ticket Count: "+str(windowsTicketCount))
print("Linux Ticket Count: "+str(linuxTicketCount))
print("Final output is: \n")
print(dfOutput.head())



dfOutput.to_excel(f'{exportPath[0]}/Cleaned combined_{int(ts)}.xlsx', index=False)
print("Operation is successful")