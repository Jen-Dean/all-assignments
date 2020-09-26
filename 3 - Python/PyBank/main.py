#Import Needed Modules
import os
import csv
import sys

#State the Variables Needed
MonthCount = 1
NetRevenue = 0
MaxRevenue = 0
MinRevenue = 0
AvgRevenue = 0
MoMChange = 0

#Create Lists to Store the CSV Data in Python
Month = []
Revenue = []
AvgChange = []

#join the file information to read from
budget_csv = os.path.join("PyBank/Resources/budget_data.csv")

print(budget_csv)

with open(budget_csv) as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")
    
    #Find the Header and First Row
    csv_header = next(csv_reader)
    firstrow = next(csv_reader)
    rowbefore = (int(firstrow[1]))

    Revenue.append(int(firstrow[1]))
    #Create the Loop to go through all the rows:
    for row in csv_reader:
        #Count the total rows in column 0 (MonthCount)
        MonthCount += 1
        #Add the integers to a database in Python [Revenue]
        Revenue.append(int(row[1]))
        #Add the Months to a database in Python [Month]
        Month.append(str(row[0]))
        #Get the month to month change and store it
        MoMChange = int(row[1]) - rowbefore
        AvgChange.append(MoMChange)
        rowbefore = int(row[1])

#Zipper the Data Together so you can pull the information later
Zipped = zip(Month, Revenue, AvgChange)

#Get the total, min and max of the integers in the list
NetRevenue = sum(Revenue)
AvgRevenue = round((sum(AvgChange)/len(AvgChange)), 2)
MaxRevenue = max(AvgChange)
MinRevenue = min(AvgChange)

#Format the numbers to show correct currency
FormatNet = "${:,.2f}".format(NetRevenue)
FormatAvg = "${:,.2f}".format(AvgRevenue)
FormatMax = "${:,.2f}".format(MaxRevenue)
FormatMin = "${:,.2f}".format(MinRevenue)

#find the month row and with the Max/Min and print
for row in Zipped:
    if row[2] == MaxRevenue:
        MaxRevMonth = str(f"{row[0]}, {FormatMax}")    
    if row[2] == MinRevenue:
        MinRevMonth = str(f"{row[0]}, {FormatMin}")

#Print the Results
print("Financial Analysis")
print("----------------------------")
print(f"Total Months: {MonthCount}")
print(f"Total: {FormatNet}")
print(f"Average Change: {FormatAvg}")
print(f"Greatest Increase in Profits: {MaxRevMonth}")
print(f"Greatest Decrease in Profits: {MinRevMonth}")


sys.stdout = open("PyBank/Analysis/Analysis.txt", 'w') 
print("Financial Analysis")
print("----------------------------")
print(f"Total Months: {MonthCount}")
print(f"Total: {FormatNet}")
print(f"Average Change: {FormatAvg}")
print(f"Greatest Increase in Profits: {MaxRevMonth}")
print(f"Greatest Decrease in Profits: {MinRevMonth}")
sys.stdout.close()