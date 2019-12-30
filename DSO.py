# Written by Berker Bahceci, July&August 2019
# DSO-Health report automation

import pandas as pd
import numpy as np
import openpyxl
import matplotlib.pyplot as plt
import matplotlib.backends.backend_pdf
from calendar import monthrange
import tkinter as tk
from tkinter import filedialog
import time
start = time.time()






###############################################################################
################### PART 1: ANALYSIS ##########################################
###############################################################################


# Get all the necessary file paths
root=tk.Tk()
root.lift()
root.attributes("-topmost", True)
aging_file_path = filedialog.askopenfilename(title='Select the SAP aging data')
customer_dict_file_path = filedialog.askopenfilename(title='Select customer dictionnary')
report_path = filedialog.asksaveasfilename(title='Save the report to', filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*") ))
sales_file_path = filedialog.askopenfilename(title='Select the SAP sales data')
sales_db_path = filedialog.askopenfilename(title='Select the sales database')
root.withdraw()




# Raw aging data
raw_age = pd.read_excel(aging_file_path)
raw_header = list(raw_age.columns)


# Credit limits
kredi_limiti = raw_age[[raw_header[8],raw_header[9],raw_header[19]]].copy()
kredi_limiti.drop_duplicates(inplace=True)
kredi_limiti = kredi_limiti.reset_index(drop=True)
kredi_limiti = kredi_limiti.dropna(axis=0, how='any')
kredi_limiti.rename(columns={raw_header[9]:'renamed'}, inplace=True)




# Grouped LOG table
# Each customer has three rows of LOG
log = raw_age[[raw_header[8], raw_header[15], raw_header[22]]].copy()
log.fillna('X', inplace=True)
log = log.groupby([raw_header[8], raw_header[22]]).sum().reset_index() 




# Initialize the summation /customer group of LOG and OAR
log_sum = pd.DataFrame(columns=['colum1', 'column2'])
log_sum['column1'] = kredi_limiti[raw_header[8]].copy()
log_sum.fillna(0, inplace=True)

risk_sum = pd.DataFrame(columns=['column1', 'column3'])
risk_sum['column1'] = kredi_limiti[raw_header[8]].copy()
risk_sum.fillna(0, inplace=True)




# Created LOG and OAR in seperate dataframes
# Get the data from those dataframes to the new, summed frames
# Reduce 3 rows of LOG to a single one
for i in range(len(log)):
    
    if log.iloc[i,1] == 'Y':
        for j in range(len(log_sum)):
            if log.iloc[i,0] == log_sum.iloc[j,0]:
                log_sum.iloc[j,1] = log.iloc[i,2]
    else:
        for j in range(len(risk_sum)):
            if log.iloc[i,0] == risk_sum.iloc[j,0]:
                risk_sum.iloc[j,1] = risk_sum.iloc[j,1]+log.iloc[i,2]




# Inıtialize dataframe to write the due date data         
vade_header = [raw_header[8], raw_header[22], raw_header[17], raw_header[18], raw_header[23], raw_header[24], raw_header[25], raw_header[26], raw_header[27], raw_header[28]]       
due_raw = raw_age[vade_header].copy()




# Drop drows with criteria='M'
# ÖDK='M' was only used in LOG calculation. They are used there. 
# They shouldn't be taken into account for OAR and due date calculations.
indexNames = due_raw[due_raw[raw_header[22]] == 'M' ].index
due_raw.drop(indexNames, inplace=True)                
due = pd.DataFrame(columns=['column1', 'column2','column3','0-30','30-60','60-90','90-120','120-150','>150'])
due['column1'] = kredi_limiti[raw_header[8]].copy()
due.fillna(0, inplace=True)


for m in range(len(due)):
    for k in range(len(due_raw)):
        if due_raw.iloc[k,0]==due.iloc[m,0]:
            due.iloc[m,1] = due.iloc[m,1]+due_raw.iloc[k,2]
            due.iloc[m,2] = due.iloc[m,2]+due_raw.iloc[k,3]
            due.iloc[m,3] = due.iloc[m,3]+due_raw.iloc[k,4]
            due.iloc[m,4] = due.iloc[m,4]+due_raw.iloc[k,5]
            due.iloc[m,5] = due.iloc[m,5]+due_raw.iloc[k,6]
            due.iloc[m,6] = due.iloc[m,6]+due_raw.iloc[k,7]
            due.iloc[m,7] = due.iloc[m,7]+due_raw.iloc[k,8]
            due.iloc[m,8] = due.iloc[m,8]+due_raw.iloc[k,9]
            



# There are negative overdue values in the data
# Take the negative overdues and add them to Not overdue, 0-30 days, without changing the sign
for m in range(len(due)):
    if due.iloc[m,1]<0:
        x = due.iloc[m,1]
        due.iloc[m,1] = 0
        due.iloc[m,2] = due.iloc[m,2]+x
        due.iloc[m,3] = due.iloc[m,3]+x




# All the data is now in the table with a single row for each customer
# Create a final dataframe with the total analysis details
final = pd.concat([kredi_limiti,log_sum['LOG'],risk_sum['Total risk'],due], axis=1)
final.drop(['some_column'], axis=1, inplace=True)
final.rename(columns={raw_header[8]:'column1'}, inplace=True)
final.loc['Total'] = final.sum(numeric_only=True, axis=0)




# Assign customers their customer groups so that group dataframes to be written into different sheets
# This part reaches to the customer dictionnary and assigns groups
final['Group']=""
customer_dictionnary = pd.read_excel(customer_dict_file_path)
names = list(customer_dictionnary.columns)
size = customer_dictionnary.shape 
for i in range(size[0]):
    for j in range(size[1]):
        for k in range(len(final)-1):
            if final.iloc[k,0] == customer_dictionnary.iloc[i,j]:
                final.iloc[k,13] = names[j]

# Create different dataframes for each customer group. 
# Each will be written in a seperate sheet.
customer1=final[final.Group == 'customer1']
customer1.loc['Total']=customer1.sum(numeric_only=True, axis=0)
customer1.loc['Total','Müsteri']='Total'
customer1.drop('Group', axis=1, inplace=True)

customer2=final[final.Group == 'customer2']
customer2.loc['Total']=customer2.sum(numeric_only=True, axis=0)
customer2.loc['Total','Müsteri']='Total'
customer2.drop('Group', axis=1, inplace=True)

customer3=final[final.Group == 'customer3']
customer3.loc['Total']=customer3.sum(numeric_only=True, axis=0)
customer3.loc['Total','Müsteri']='Total'
customer3.drop('Group', axis=1, inplace=True)

customer4=final[final.Group == 'customer4']
customer4.drop('Group', axis=1, inplace=True)

customer5=final[final.Group == 'customer5']
customer5.drop('Group', axis=1, inplace=True)

customer6=final[final.Group == 'customer6']
customer6.drop('Group', axis=1, inplace=True)

customer7=final[final.Group == 'customer7']
customer7.drop('Group', axis=1, inplace=True)

customer8=final[final.Group == 'customer8']
customer8.loc['Total']=customer8.sum(numeric_only=True, axis=0)
customer8.loc['Total','Müsteri']='Total'
customer8.drop('Group', axis=1, inplace=True)

customer9=final[final.Group == 'customer9']
customer9.loc['Total']=customer9.sum(numeric_only=True, axis=0)
customer9.loc['Total','Müsteri']='Total'
customer9.drop('Group', axis=1, inplace=True)

customer10=final[final.Group == 'customer10']
customer10.loc['Total']=customer10.sum(numeric_only=True, axis=0)
customer10.loc['Total','Müsteri']='Total'
customer10.drop('Group', axis=1, inplace=True)


final.drop('Group', axis=1, inplace=True)
customers = [customer1, customer2, customer3, customer4, customer5, customer6, customer7, customer8, customer9, customer10]



# The report template gets written
# Each customer has a sheet with the LOG, Credit Limit, OAR and Due data
with pd.ExcelWriter('%s.xlsx' % report_path) as writer:
    customer1.to_excel(writer, sheet_name='customer1', index=False, header=final.keys())
    customer2.to_excel(writer, sheet_name='customer2', index=False, header=final.keys())
    customer3.to_excel(writer, sheet_name='customer3', index=False, header=final.keys())
    customer4.to_excel(writer, sheet_name='customer4', index=False, header=final.keys())
    customer5.to_excel(writer, sheet_name='customer5', index=False, header=final.keys())
    customer6.to_excel(writer, sheet_name='customer6', index=False, header=final.keys())
    customer7.to_excel(writer, sheet_name='customer7', index=False, header=final.keys())
    customer8.to_excel(writer, sheet_name='customer8', index=False, header=final.keys())
    customer9.to_excel(writer, sheet_name='customer9', index=False, header=final.keys())
    customer10.to_excel(writer, sheet_name='customer10', index=False, header=final.keys())
   






###############################################################################
################### PART 2: SALES #############################################
###############################################################################
    

# Get the Excel file which has the monthly sales data
raw_sales = pd.read_excel(sales_file_path)
raw_sales = raw_sales.drop(len(raw_sales)-1, axis=0)




# Remove unnecessary columns&calculate sum
raw_sales_header = list(raw_sales.columns)
sales_header = [raw_sales_header[8], raw_sales_header[13], raw_sales_header[14], raw_sales_header[2]]
sales = raw_sales[sales_header].copy()
grouped_sales = sales.groupby(['some_column','some_column2','some_column3']).sum().reset_index()
grouped_sales['OAR'] = ""




# Assign customers their customer groups so that group dataframes to be written into different sheets
# This part reaches to the customer dictionnary and assigns groups
group_names = list(customer_dictionnary.columns)
grouped_sales['Group'] = ""

size = customer_dictionnary.shape 
for i in range(size[0]):
    for j in range(size[1]):
        for k in range(len(grouped_sales)):
           if grouped_sales.iloc[k,1] == customer_dictionnary.iloc[i,j]:
                grouped_sales.iloc[k,5] = group_names[j]




# Add a totals row for cumulative sales
# This total cumulative sales will be written to update the Sales,OAR&DSO database
grouped_sales.loc['Total'] = grouped_sales.sum(numeric_only=True, axis=0)
grouped_sales.loc['Total','some_column'] = final.loc['Total','some_other_column']
grouped_sales.loc['Total','some_column7'] = 'Total'
grouped_sales.rename(columns={'some_column':'some_column_new'}, inplace=True)
grouped_sales.loc['Total','Group'] = 'Total'







##############################################################################
################### PART 3: UPDATE THE REPORT TEMPLATE AND DATABASES #########
##############################################################################


# Update the database file with this month's Sales&OAR
wb = openpyxl.load_workbook(sales_db_path)
ws = wb.active
max_columns = ws.max_column
max_rows = ws.max_row
while ws.cell(row=max_rows, column=2).value == None:
        ws.delete_rows(max_rows,1)
        max_rows = ws.max_row




for i in range(2, max_columns, 3):
   ws.cell(row = max_rows + 1, column = i).value = dso_sales_final.iloc[int(np.floor(i/3)),0]

for i in range(3, max_columns + 1, 3):
   ws.cell(row = max_rows + 1, column = i).value = dso_sales_final.iloc[int(i/3)-1,1]

ws.cell(row = max_rows + 1, column = 1).value=grouped_sales.iloc[0,1]
     
max_columns = ws.max_column
max_rows = ws.max_row
for i in range(1, max_rows+5):
    if ws.cell(row=i, column=2).value == None:
        ws.delete_rows(i,1)
wb.save(sales_db_path)




# Load the updated sales databse and calculate the DSO
dso_table = pd.read_excel(sales_db_path)  #If it gives an error here, use %s %
length_dso = len(dso_table)
for i in range(length_dso-12):
    dso_table.drop(i, axis=0, inplace=True)
    
dso_table = dso_table.reset_index()
dso_table.drop('index', axis=1, inplace=True)


# DSO calculator
for i in range (3, max_columns,3):
    c = 0
    days = 0
    dso = dso_table.iloc[11,i-1]
    for j in range(11,5,-1):
        if dso-dso_table.iloc[j,i-2]>0:
            dso = dso-dso_table.iloc[j,i-2]
            days = days+monthrange(dso_table.iloc[j,0].year, dso_table.iloc[j,0].month)[1]
            c = c+1
        else: 
            break
    
    dso_table.iloc[11,i] = (dso/dso_table.iloc[j,i-2])*30+days


# Write calculated DSO value back into sales database.    
wb = openpyxl.load_workbook(sales_db_path)
ws = wb.active
max_row = ws.max_row
for i in range(3,max_columns,3):
    ws.cell(row=max_row, column=i+1).value = dso_table.iloc[len(dso_table)-1, i]
wb.save(sales_db_path)





# Update the report template with the DSO value for this month which is calculated above
i=0;j=0
wb = openpyxl.load_workbook('%s.xlsx' %report_path)
sheets = wb.sheetnames
for i in range(len(sheets)):
    ws = wb[sheets[i]]
    max_rows = ws.max_row
    max_columns = ws.max_column
    for j in range(len(dso_table)):
        #Give context to the DSO table with this.
        ws.cell(row=1, column=max_columns+4).value = 'DSO Table'
        ws.cell(row=2, column=max_columns+4).value = 'Monthly sales'
        ws.cell(row=2, column=max_columns+5).value = 'Open account risk'
        ws.cell(row=2, column=max_columns+6).value = 'DSO'
        ws.cell(row=2, column=max_columns+7).value = 'Period'
        
        ws.cell(row=j+3, column=max_columns+4).value = dso_table.iloc[j,3*(i+1)-2]
        ws.cell(row=j+3, column=max_columns+5).value = dso_table.iloc[j,3*(i+1)-1]
        ws.cell(row=j+3, column=max_columns+6).value = dso_table.iloc[j,3*(i+1)]
        ws.cell(row=j+3, column=max_columns+7).value = dso_table.iloc[j,0]



wb.save('%s.xlsx' % report_path)





##############################################################################
################### PART 4: FINALIZE REPORT WITH CHARTS ######################
##############################################################################

# Fixed values to be used in each chart is calculated before the loop for efficiency
wb = openpyxl.load_workbook('%s.xlsx' % report_path)
sheets = wb.sheetnames
all_customer_dso=dso_table['Unnamed: 33']
period=dso_table['Unnamed: 0'].dt.to_period('M')
credit_lim_total = final.iloc[len(final)-1, 2]   
ar_total = final.iloc[len(final)-1, 4]           
pos = list(range(len(all_customer_dso)))
width = 0.25
min1 = min(all_customer_dso)
max1 = max(all_customer_dso)
pdf = matplotlib.backends.backend_pdf.PdfPages("Charts PDF output.pdf")


for i in range(0, len(sheets)):
    if sheets[i]=='All customers':
        continue
    
    ws = wb[sheets[i]]
    max_rows = ws.max_row
    max_columns = ws.max_column




    # Sales pie
    # Labels and values for the chart
    labels = [sheets[i], 'Rest of the sales']
    x = dso_sales_final.iloc[i,0]/dso_sales_final.loc['Total', 'Tutar']
    sizes = [x, 1-x]
    
    explode = (0.1, 0) 
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, colors=['tan', 'lightblue'], autopct='%1.1f%%', startangle=15)
    ax1.axis('equal')
    plt.figure(figsize=(4,2))
    
    
    # Title and legend
    ax1.set_title("%s's Monthly Sales Volume" %sheets[i])
    ax1.legend(labels, loc='best', bbox_to_anchor=(0.3, 0.2,))
    fig1.savefig("sales_pie%s.png" % i, dpi = 72)
    pdf.savefig(fig1)


    

    # DSO bar
    fig2, ax2 = plt.subplots(figsize=(14,8))
    # Insert data into bars
    plt.bar(pos, all_customer_dso,width, label='All customers', color='lightblue')
    plt.bar([p + width for p in pos], dso_table.iloc[:,3*i+3], width, label='%s' % sheets[i], color='tan')


    # Labels
    ax2.set_ylabel('DSO')
    ax2.set_title("%s's Yearly DSO Performance in Comparison with Total DSO" %sheets[i])
    ax2.set_xticks([p+0.1 for p in pos])
    ax2.set_xticklabels(period)
    
    
    # Set the range of the y-axis for better usage
    plt.ylim(60 , 130)
    
    
    plt.grid()
    plt.legend(['Total', '%s' % sheets[i]])
    fig2.savefig("bar%s.png" %i, dpi=72)
    pdf.savefig(fig2)
    
    
    
    
    # Risk pie
    # Labels and values for the chart
    labels3 = [sheets[i], 'Rest of the A/R']
    x3 = customers[i].iloc[len(customers[i]) -1, 4]/ar_total
    sizes3 = [x3, 1-x3]
    
    explode = (0.1, 0) 
    fig3, ax3 = plt.subplots()
    ax3.pie(sizes3, explode=explode, colors=['tan', 'lightblue'], autopct='%1.1f%%', startangle=15)
    ax3.axis('equal')
    plt.figure(figsize=(4,2))
    
    
    # Title and legend
    ax3.set_title("%s's A/R ratio" %sheets[i])
    ax3.legend(labels3, loc='best', bbox_to_anchor=(0.3, 0.2,))
    fig3.savefig("risk_pie%s.png" % i, dpi = 72)
    pdf.savefig(fig3)
    
    
    
    
    # LoG pie
    # Labels and values for the chart
    labels4 = [sheets[i], 'Rest of the credit limit']
    x4 = customers[i].iloc[len(customers[i]) -1, 2]/credit_lim_total
    sizes4 = [x4, 1-x4]
    
    explode = (0.1, 0) 
    fig4, ax4 = plt.subplots()
    ax4.pie(sizes4, explode=explode, colors=['tan', 'lightblue'], autopct='%1.1f%%', startangle=15)
    ax4.axis('equal')
    plt.figure(figsize=(4,2))
    
    
    # Title and legend
    ax4.set_title("%s's Credit Limit Ratio" %sheets[i])
    ax4.legend(labels4, loc='best', bbox_to_anchor=(0.3, 0.2,))
    fig4.savefig("cl_pie%s.png" % i, dpi = 72)
    pdf.savefig(fig4)
    
    
    
    
    # Add to Excel file
    img1 = openpyxl.drawing.image.Image('sales_pie%s.png' % i)
    ws.add_image(img1, ws.cell(row=max_rows+4, column=1).coordinate)
    img2 = openpyxl.drawing.image.Image('bar%s.png' %i)
    ws.add_image(img2, ws.cell(row=16, column=16).coordinate)
    img3 = openpyxl.drawing.image.Image('risk_pie%s.png' %i)
    ws.add_image(img3, ws.cell(row=max_rows+4, column=7).coordinate)   
    img4 = openpyxl.drawing.image.Image('cl_pie%s.png' %i)
    ws.add_image(img4, ws.cell(row=max_rows+20, column=1).coordinate)   
    
    
wb.save('%s.xlsx' % report_path)
pdf.close()




# Measure the time elapsed from the start
end = time.time()
duration = end-start
print('Report is generated in %s seconds' %duration)
################################ END ##########################################
