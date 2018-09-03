

from openpyxl import *
from openpyxl.styles import colors,Font

def applyStyles () :
    output_wb = load_workbook("output.xlsx")
    work_sheets= output_wb.worksheets
    current_worksheet = work_sheets[0]
    paddy_purchases = current_worksheet["D12"].value + current_worksheet["D14"].value + current_worksheet["D15"].value + current_worksheet["D17"].value

    for row in current_worksheet : 
        for cell in row :
            if cell.col_idx == 10 :
                if isinstance(cell.value,(int,float,long)) :
                    next_cell_id = "K" +str(cell.row)
                    print "Next Cell Value  "+str(current_worksheet[next_cell_id].value)
                    if current_worksheet[next_cell_id].value =="Error"  :
                    #print "Row  "+ str(cell.row)
                    #print "Column  "+str(cell.col_idx)
                        cell_id = "J"+str(cell.row)
                        print next_cell_id
                        style_cell= current_worksheet[cell_id]
                        ft = Font(color=colors.RED)
                        style_cell.font = ft  

                    else :
                        cell_id = "J"+str(cell.row)
                        print cell_id
                        style_cell= current_worksheet[cell_id]
                        ft = Font(color=colors.GREEN)
                        style_cell.font = ft  

            if cell.col_idx == 12 :
                if isinstance(cell.value,(int,float,long)) :
                    next_cell_id = "M" +str(cell.row)
                    print "Next Cell Value  "+str(current_worksheet[next_cell_id].value)

                    if current_worksheet[next_cell_id].value =="Error"  :
                    #print "Row  "+ str(cell.row)
                    #print "Column  "+str(cell.col_idx)
                        cell_id = "L"+str(cell.row)
                        print cell_id
                        style_cell= current_worksheet[cell_id]
                        ft = Font(color=colors.RED)
                        style_cell.font = ft  

                    else :
                        cell_id = "L"+str(cell.row)
                        print cell_id
                        style_cell= current_worksheet[cell_id]
                        ft = Font(color=colors.GREEN)
                        style_cell.font = ft  

    current_worksheet = work_sheets[1]
    #Second Worksheet
    
    paddy_milling = None
    paddy_milling_cell = current_worksheet["F15"]
    paddy_milling = paddy_milling_cell.value
    print "Paddy Milling  "+str(paddy_milling)
   
    for row in current_worksheet : 
        for cell in row :
            if cell.row == 47 and cell.col_idx ==6 :
                if cell.value > 69 and cell.value <70 :
                    cell_id = "F"+str(cell.row)
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.GREEN)
                    style_cell.font = ft  
                else :
                    cell_id = "F"+str(cell.row)
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.RED)
                    style_cell.font = ft  


    #Third Work Sheet
    current_worksheet = work_sheets[2]
    electricity_charges = None
    electricity_charges_cell = current_worksheet["C11"]
    electricity_charges = electricity_charges_cell.value
    print "Electricity Charges  "+str(electricity_charges)

    print "Paddy Milling * 60   " +str(paddy_milling*60)
    print "Paddy Milling *75  " +str(paddy_milling*75)
    
    hamali_charges = None
    hamali_charges = current_worksheet["C28"].value
    hamali_charges_flag = None

    if hamali_charges >(paddy_purchases * 35)/100 :
        hamali_charges_flag = True
        cell_id = "C28"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.RED)
        style_cell.font = ft  
        print "Hamali Charges Error"
    else :
        hamali_charges_flag = False
        cell_id = "C28"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.GREEN)
        style_cell.font = ft  
        print "Hamali Charges Correct"
    

    labour_charges = current_worksheet["C29"].value
    labour_charges_cell = current_worksheet["C29"]
    
    if (paddy_purchases *4)/100 > labour_charges :
        hamali_charges_flag = True
        cell_id = "C29"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.RED)
        style_cell.font = ft  
        print "labour Charges Error"
    else :
        cell_id = "C29"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.GREEN)
        style_cell.font = ft  
        print "Labour Charges Correct"


    if paddy_milling * 60 < electricity_charges and paddy_milling*75 > electricity_charges :
        print "paddy Milling Correct"
        cell_id = "C11"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.RED)
        style_cell.font = ft  

    else :
        print "paddy Milling Error"
        cell_id = "C11"
        print cell_id
        style_cell= current_worksheet[cell_id]
        ft = Font(color=colors.RED)
        style_cell.font = ft  

    current_worksheet = work_sheets[3]
    prev_cell_value = None
    trade_creditors = None
    sunday_debtors = None
    
    

    for row in current_worksheet : 
        for cell in row :
        
            if cell.row == 24 and cell.col_idx == 3 :
                if paddy_milling /4 > trade_creditors :
                    cell_id = "B24"
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.RED)
                    style_cell.font = ft  
                else :
                    cell_id = "B24"
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.GREEN)
                    style_cell.font = ft  
            elif cell.row == 24 and cell.col_idx == 5:
                if cell.value > (paddy_purchases*15)/100 :
                    cell_id = "E24"
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.GREEN)
                    style_cell.font = ft 
                else :
                    cell_id = "E24"
                    print cell_id
                    style_cell= current_worksheet[cell_id]
                    ft = Font(color=colors.GREEN)
                    style_cell.font = ft 
            if prev_cell_value == "Trade Creditors" :
                trade_creditors = cell.value
            if prev_cell_value == "Sundry Debtors" :
                sunday_debtors = cell.value
            prev_cell_value = cell.value
    
    output_wb.save("output.xlsx")
    output_wb.close()