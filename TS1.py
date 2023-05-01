

#%%
#%%
import pandas as pd
import numpy as np
import random
import math
import os
import datetime
from random import choice
from more_itertools import locate
from collections import OrderedDict
import time
import logging
import matplotlib.pyplot as plt
import sys
import copy

#%%
# read data from input file and return parameters/variables
def read_data(df_VM_info, input_excel, input_sheet_ProEast, input_sheet_ProNotEast, input_sheet_ProDemand, Index_strart, Index_end):  
    
    # Input data from (VM_info.sheet)
    # CargoLane_TotalNumber = int(df_VM_info.at[Index_strart, "CargoLane_TotalNumber"])
    # CargoLane_TotalNumber = int(len((df_VM_info["CargoLane_ID"].squeeze()).tolist()) -1) # 直接抓最大的ID
    CargoLane_TotalNumber = int((df_VM_info["CargoLane_TotalNumber"].squeeze()).tolist()[0])
    VM_ID = df_VM_info.at[Index_strart, "VM_ID"]
    CargoLane_Device_ID = (df_VM_info.loc[Index_strart:Index_end, ["Device_ID"]].squeeze()).tolist()
    CargoLane_Site_ID = (df_VM_info.loc[Index_strart:Index_end, ["Site_ID"]].squeeze()).tolist()
    CargoLane_ID = (df_VM_info.loc[Index_strart:Index_end, ["CargoLane_ID"]].squeeze()).tolist()
    CargoLane_Type = (df_VM_info.loc[Index_strart:Index_end, ["CargoLane_Type"]].squeeze()).tolist()
    CargoLane_Height_Max = (df_VM_info.loc[Index_strart:Index_end, ["High_Max"]].squeeze()).tolist()
    CargoLane_Height_Min = (df_VM_info.loc[Index_strart:Index_end, ["High_Min"]].squeeze()).tolist()
    CargoLane_Diameter_Max_1 = (df_VM_info.loc[Index_strart:Index_end, ["Diameter_Max_1"]].squeeze()).tolist()
    CargoLane_Diameter_Min_1 = (df_VM_info.loc[Index_strart:Index_end, ["Diameter_Min_1"]].squeeze()).tolist()
    CargoLane_Area = (df_VM_info.loc[Index_strart:Index_end, ["Area"]].squeeze()).tolist()
    CargoLane_Capacity = (df_VM_info.loc[Index_strart:Index_end, ["CargoLane_Capacity"]].squeeze()).tolist()
    Current_Product = (df_VM_info.loc[Index_strart:Index_end, ["Current_Product"]].squeeze()).tolist()
    Max_Prod_Cnt = df_VM_info.at[Index_strart, "Max_Prod_Cnt"]
    Min_Prod_Cnt = df_VM_info.at[Index_strart, "Min_Prod_Cnt"]
    CargoLane_Allow_Special = (df_VM_info.loc[Index_strart:Index_end, ["Allow_Special"]].squeeze()).tolist()
    CargoLane_Average_Replenishment = (df_VM_info.loc[Index_strart:Index_end, ["Average_Replenishment"]].squeeze()).tolist()
    CargoLane_Category_Rate = (df_VM_info.loc[Index_strart:Index_end, ["Category_Rate"]].squeeze()).tolist()
    CargoLane_Brand_Rate = (df_VM_info.loc[Index_strart:Index_end, ["Brand_Rate"]].squeeze()).tolist()
    
    # distinguish the cargolane type which can allow special size from normal(can not allow)
    #for i in range(len(CargoLane_Type)):
        #if CargoLane_Type[i] == 5:
            #pass
        #elif CargoLane_Allow_Special[i] == 1:
            #CargoLane_Type[i] = "s" + str(CargoLane_Type[i])
            
        #CargoLane_ID[i] = int(CargoLane_ID[i]) # !!!!!
    for i in range(len(CargoLane_Type)): 
        CargoLane_Type[i] = int(CargoLane_Type[i])
        CargoLane_ID[i] = int(CargoLane_ID[i])
        
  
    df_Product_info = pd.read_excel(input_excel, sheet_name = "Product_info")
    
    # Input data from (Product_info.sheet)
    Product_ID = df_Product_info["Product_ID"].tolist()
    Product_Price = df_Product_info["Price"].tolist()
    Product_Cost = df_Product_info["Cost"].tolist()
    Product_Product_sales = df_Product_info["Average_sales_month"].tolist()
    Product_Type = df_Product_info["Type"].tolist()
    Product_Volume = df_Product_info["Volume"].tolist()
    Product_Length = df_Product_info["length"].tolist()
    Product_Width = df_Product_info["width"].tolist()
    Product_Height = df_Product_info["height"].tolist()
    Product_New = df_Product_info["New"].tolist()
    Product_Brand = df_Product_info["Brand"].tolist()
    Product_Category = df_Product_info["Category"].tolist()
    Product_Specialsize = df_Product_info["Special_size"].tolist()
    total_cost = df_Product_info["Total_Cost"].tolist()
    #Product_Operating_cost = df_Product_info["Operating_cost"].tolist()
    #Product_Total_produced = df_Product_info["Total_produced"].tolist()
   
    # unit_purc_cost=[]
    #pr_price=[]
    '''
    for num1,num2 in zip(Product_Operating_cost, Product_Total_produced):
        unit_purc_cost.append(num1/num2) # Operating Cost/ Total Produced
        #unit_purc_cost.append(num1+num2) # Operating Cost/ Total Produced
    '''
    # random.seed(0)
    # for i in range(len( Product_ID)):
    #     rand_purch = random.uniform(10, 20)
    #     #rand_pr= random.uniform(20, 40)
    #     unit_purc_cost.append(rand_purch)
    #     #pr_price.append(rand_pr)
    # Product_Cost= unit_purc_cost
    #Product_Price= pr_price
    #print("Product_Cost=", Product_Cost)
    #print("Product_Price=", Product_Price)
   
    df_Product_demand = pd.read_excel(input_excel, sheet_name = input_sheet_ProDemand)
    
    # Input data from (Product_demand.sheet)
    Demand_Product_ID = df_Product_demand["Product_ID"].tolist()
    Demand_Product_Sales = df_Product_demand["Average_sales_month"].tolist()
    
    # variable for saving ID which Demand_sales = 0
    Demand_zero = []
    for i in range(len(Demand_Product_Sales)):
        if Demand_Product_Sales[i] == 0:
            Demand_zero.append(Demand_Product_ID[i])
    
    # Input data from (Product_Repel.sheet)
    df_replacement_matrix = pd.read_excel(input_excel, sheet_name = "Product_Repel")
    df_replacement_matrix.set_index("Unnamed: 0", inplace=True)
    
    replacement_index = df_replacement_matrix.index
    replacement_matrix = {}
    for i in range(len(replacement_index)):
        # print(replacement_index[i])
        # print(df_replacement_matrix[replacement_index[i]].loc[df_replacement_matrix[replacement_index[i]]==1].keys().tolist()[0])
        replacement_matrix.setdefault(replacement_index[i], df_replacement_matrix[replacement_index[i]].loc[df_replacement_matrix[replacement_index[i]]==1].keys().tolist()[0])
     
    # 0913
    for i in range(len(Product_New)):
         if Product_New[i] == 1 and Product_ID[i] in Demand_Product_ID:
             Product_New[i] = 0
             
    return total_cost, df_VM_info, df_Product_info, df_Product_demand, df_replacement_matrix, VM_ID, CargoLane_Device_ID, CargoLane_Site_ID, CargoLane_TotalNumber, CargoLane_ID, CargoLane_Type, \
        CargoLane_Height_Max, CargoLane_Height_Min, CargoLane_Diameter_Max_1, CargoLane_Diameter_Min_1, \
        CargoLane_Area, CargoLane_Capacity, Current_Product, Max_Prod_Cnt, Min_Prod_Cnt, CargoLane_Allow_Special, \
        CargoLane_Average_Replenishment, CargoLane_Category_Rate, CargoLane_Brand_Rate, \
        Product_ID, Product_Price, Product_Cost, Product_Product_sales, Product_Type, Product_Volume, Product_Length, Product_Width, Product_Height, Product_New, \
        Product_Brand, Product_Category, Product_Specialsize, Demand_Product_ID, Demand_Product_Sales, replacement_matrix, Demand_zero


#%%
def classify_demand_product(Product_ID, Product_Type, Product_Volume, Product_Price, Demand_Product_ID, Demand_Product_Sales, CargoLane_Average_Replenishment, Product_New, Product_Brand):

    ID_CargoLane1 = []
    ID_CargoLane2 = []
    ID_CargoLane3 = []
    ID_CargoLane4 = []
    ID_CargoLane5 = []
    
    Price_CargoLane1 = []    
    Price_CargoLane2 = []    
    Price_CargoLane3 = []    
    Price_CargoLane4 = []    
    Price_CargoLane5 = []

    Sales_CargoLane1 = []    
    Sales_CargoLane2 = []    
    Sales_CargoLane3 = []    
    Sales_CargoLane4 = []    
    Sales_CargoLane5 = []
    
    replenishment_per_time = []                                                # demand/replenishment lead time
    Cargolanetype_sum_capacity = [0] * 6                                       # sum of the capacity with same cargolane type 
    Cargolanetype_average_capacity = [0] * 6                                   # average of the capacity with same cargolane type
    product_typenum = [6] * len(Demand_Product_ID)                             # each product type
    Product_max_cargolanenum = []
    
    New_ID1 = []
    New_ID2 = []
    New_ID3 = []
    New_ID4 = []
    New_ID5 = []
    
    Brand_CargoLane1 = []
    Brand_CargoLane2 = []
    Brand_CargoLane3 = []
    Brand_CargoLane4 = []
    Brand_CargoLane5 = []
    
    cargolane_should_empty = []
    
    # inventory_cost=[]
    # backroom_cost=[]
    # display_cost=[]
    # ordering_cost=[]
    # total_cost=[]
    setup_cost=[]
    replenishment_cost=[]
    
    # calculate cost
    
    #Inventory Cost
    # random.seed(0)
    # for i in range(len(Product_ID)):
    #     rand_inv = random.uniform(2.0, 4.0)
    #     rand_backroom= random.uniform(1.0, 2.0)
    #     rand_display= random.uniform(1.0, 3.0)
    #     inventory_cost.append(rand_inv)
    #     backroom_cost.append(rand_backroom)
    #     display_cost.append(rand_display)
    #print("Inv_Cost=",inventory_cost)
    #print("Backroom_Cost=",backroom_cost)
    #print("Display_Cost=",display_cost)
    
    # for i in Product_Price:
        # ordering_cost.append(50) #0
    #print(ordering_cost)   
    '''
      # assume capacity =15
    capacity=15
    
    for i in Product_Cost:
        inventory_cost.append(i*capacity) #purchase cost* capacity
        #inventory_cost.append(i*0) #purchase cost* capacity
        
        
    for i in Product_Price:
        
        backroom_cost.append(i*0.5) #price* 5%
        display_cost.append(i*0.7* capacity) #price* 7% * capacity
        ordering_cost.append(i*0) #0
        
        #backroom_cost.append(i*0) #price* 5%
        #display_cost.append(i*0) #price* 7% * capacity
        #ordering_cost.append(i*0) #0
    
    
    '''
    # for num1,num2,num3,num4 in zip(inventory_cost,backroom_cost,display_cost,ordering_cost):
    #     total_cost.append(num1+num2+num3+num4) #the total cost includes inventory_cost,backroom_cost,display_cost,ordering_cost for layer 2
    # print("Total_Cost=",total_cost)
    
    count= CargoLane_Capacity.count(0)
    count_cap= len(CargoLane_Capacity) - count
    
    setup= 0
    replenishment_time= 0
    rep_fee= 0
    replenishment= rep_fee * replenishment_time
    
    setupcost= setup/count_cap
    replenishmentcost= replenishment/count_cap
    
    for i in range(len(Product_ID)):
        setup_cost.append(setupcost)
        replenishment_cost.append(replenishmentcost)
    
    #setup_cost= 0
    #replenishment_cost= 0
    
    '''
    print(Product_Cost)
    print(inventory_cost)
    
    print(Product_Price)
    print(inventory_cost)
    print(backroom_cost)
    print(display_cost)
    print(ordering_cost)
    print(total_cost)
    print(count)
    print(count_cap)
    print(setup_cost)
    print(replenishment_cost)
    '''
    
    #print("setup cost=",setup_cost)
    #print("replenishment cost=",replenishment_cost)
   
    for i in range(0, len(Demand_Product_ID), 1):
        replenishment_per_time.append((Demand_Product_Sales[i] / (1440 * 30 / (CargoLane_Average_Replenishment[0] * 1440))))
        for j in range(0, len(Product_ID), 1):
            if Demand_Product_ID[i] == Product_ID[j]:
                
                if Product_Type[j] == "CAN" and Product_Volume[j] <= 330:
                    ID_CargoLane1.append(Demand_Product_ID[i])
                    Price_CargoLane1.append(Product_Price[j])
                    Sales_CargoLane1.append(Demand_Product_Sales[i])
                    Brand_CargoLane1.append(Product_Brand[j])
                    if product_typenum[i] > 1:
                        product_typenum[i] = 1
                    if Product_New[j] == 1:
                        New_ID1.append(Product_ID[j])
                        
                if Product_Type[j] == "CAN" and Product_Volume[j] <= 330 \
                    or Product_Type[j] == "SCAN" and Product_Volume[j] <= 330:
                    ID_CargoLane2.append(Demand_Product_ID[i])
                    Price_CargoLane2.append(Product_Price[j])
                    Sales_CargoLane2.append(Demand_Product_Sales[i])
                    Brand_CargoLane2.append(Product_Brand[j])
                    if product_typenum[i] > 2:
                        product_typenum[i] = 2
                    if Product_New[j] == 2:
                        New_ID2.append(Product_ID[j])
                
                if Product_Type[j] == "PET" and Product_Volume[j] <= 500:
                    ID_CargoLane3.append(Demand_Product_ID[i])
                    Price_CargoLane3.append(Product_Price[j])
                    Sales_CargoLane3.append(Demand_Product_Sales[i])
                    Brand_CargoLane3.append(Product_Brand[j])
                    if product_typenum[i] > 3:
                        product_typenum[i] = 3
                    if Product_New[j] == 3:
                        New_ID3.append(Product_ID[j])
                    
                if Product_Type[j] == "PET" and Product_Volume[j] <= 600:
                    ID_CargoLane4.append(Demand_Product_ID[i])
                    Price_CargoLane4.append(Product_Price[j])
                    Sales_CargoLane4.append(Demand_Product_Sales[i])
                    Brand_CargoLane4.append(Product_Brand[j])
                    if product_typenum[i] > 4:
                        product_typenum[i] = 4
                    if Product_New[j] == 4:
                        New_ID4.append(Product_ID[j])
                    
                if Product_Type[j] == "PET" and Product_Volume[j] <= 600:
                    ID_CargoLane5.append(Demand_Product_ID[i])
                    Price_CargoLane5.append(Product_Price[j])
                    Sales_CargoLane5.append(Demand_Product_Sales[i])
                    Brand_CargoLane5.append(Product_Brand[j])
                    if product_typenum[i] > 5:
                        product_typenum[i] = 5
                    if Product_New[j] == 5:
                        New_ID5.append(Product_ID[j])
    '''
    for i in range(len(CargoLane_Type)):
        if CargoLane_Type[i] == 1:
            Cargolanetype_sum_capacity[1] += CargoLane_Capacity[i]
        elif CargoLane_Type[i] == 2:
            Cargolanetype_sum_capacity[2] += CargoLane_Capacity[i]
        elif CargoLane_Type[i] == 3:
            Cargolanetype_sum_capacity[3] += CargoLane_Capacity[i]
        elif CargoLane_Type[i] == 4:
            Cargolanetype_sum_capacity[4] += CargoLane_Capacity[i]
        elif CargoLane_Type[i] == 5:
            Cargolanetype_sum_capacity[5] += CargoLane_Capacity[i]
        else:
            cargolane_should_empty.append(i)                                   # cargolane index!!!!
    
    count_type_num = 0
    
    for j in range(1, len(Cargolanetype_average_capacity)):                    # 計算Cargolane type的平均容量, 用於計算商品最大貨道數
        if CargoLane_Type.count(j) == 0:
            Cargolanetype_average_capacity[j] = 0
        elif j == 1 or j ==2:
            count_type_num += CargoLane_Type.count(j)
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[:j+1]) / count_type_num
        elif j == 3:
            count_type_num = 0
            count_type_num += CargoLane_Type.count(j)
            Cargolanetype_average_capacity[j] = Cargolanetype_sum_capacity[j] / count_type_num
        else:
            count_type_num += CargoLane_Type.count(j)
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[3:j+1]) / count_type_num
    print(Cargolanetype_average_capacity)
    '''
   
    # calculate the numbers and capacity of every cargolane type(0~5)
    for i in range(CargoLane_TotalNumber):                                     
        if type(CargoLane_Type[i]) == float or type(CargoLane_Type[i]) == int:
            if CargoLane_Type[i] == 1:
                Cargolanetype_sum_capacity[1] += CargoLane_Capacity[i]
            elif CargoLane_Type[i] == 2:
                Cargolanetype_sum_capacity[2] += CargoLane_Capacity[i]
            elif CargoLane_Type[i] == 3:
                Cargolanetype_sum_capacity[3] += CargoLane_Capacity[i]
            elif CargoLane_Type[i] == 4:
                Cargolanetype_sum_capacity[4] += CargoLane_Capacity[i]
            elif CargoLane_Type[i] == 5:
                Cargolanetype_sum_capacity[5] += CargoLane_Capacity[i]
            else:
                cargolane_should_empty.append(i)                               # cargolane index!!!!
        else:
            if CargoLane_Type[i][1] == "1":
                Cargolanetype_sum_capacity[1] += CargoLane_Capacity[i]
            elif CargoLane_Type[i][1] == "2":
                Cargolanetype_sum_capacity[2] += CargoLane_Capacity[i]
            elif CargoLane_Type[i][1] == "3":
                Cargolanetype_sum_capacity[3] += CargoLane_Capacity[i]
            elif CargoLane_Type[i][1] == "4":
                Cargolanetype_sum_capacity[4] += CargoLane_Capacity[i]
            elif CargoLane_Type[i][1] == "5":
                Cargolanetype_sum_capacity[5] += CargoLane_Capacity[i]
            else:
                cargolane_should_empty.append(i)                               # cargolane index!!!!
    
    count_type_num = 0
    
    # calculate the average capacity of different cargolane type, it's for calculating the max number of each product
    for j in range(1, len(Cargolanetype_average_capacity)):                    
        if CargoLane_Type.count(j) == 0:
            Cargolanetype_average_capacity[j] = 0
        elif j == 1:
            count_type_num += CargoLane_Type.count(j)
            #count_type_num += CargoLane_Type.count("s1.0")
            #count_type_num += CargoLane_Type.count("s1")
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[:j+1]) / count_type_num
        elif j == 2:
            count_type_num += CargoLane_Type.count(j)
            #count_type_num += CargoLane_Type.count("s2.0")
            #count_type_num += CargoLane_Type.count("s2")
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[:j+1]) / count_type_num
        elif j == 3:
            count_type_num = 0
            count_type_num += CargoLane_Type.count(j)
            #count_type_num += CargoLane_Type.count("s3.0")
            #count_type_num += CargoLane_Type.count("s3")
            Cargolanetype_average_capacity[j] = Cargolanetype_sum_capacity[j] / count_type_num
        elif j == 4:
            count_type_num += CargoLane_Type.count(j)
            #count_type_num += CargoLane_Type.count("s4.0")
            #count_type_num += CargoLane_Type.count("s4")
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[3:j+1]) / count_type_num
        else:
            count_type_num += CargoLane_Type.count(j)
            Cargolanetype_average_capacity[j] = sum(Cargolanetype_sum_capacity[3:j+1]) / count_type_num
    
    #print(Cargolanetype_average_capacity)
    
    for i in range(len(Demand_Product_Sales)):                                 # 判斷商品類別
        if product_typenum[i] == 6:
            Product_max_cargolanenum.append(0)
        elif product_typenum[i] == 1:
            type_cap = Cargolanetype_average_capacity[2]
            if type_cap != 0:
                Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] / Cargolanetype_average_capacity[2]))
            else:
                type_cap = Cargolanetype_average_capacity[product_typenum[1]]
                if type_cap != 0:
                    Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] / Cargolanetype_average_capacity[1]))
                else:
                    Product_max_cargolanenum.append(0)
                    
        elif product_typenum[i] == 2:
            type_cap = Cargolanetype_average_capacity[2]
            if type_cap != 0:
                Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] /  Cargolanetype_average_capacity[2]))
            else:
                Product_max_cargolanenum.append(0)
                
        elif product_typenum[i] == 3:
            type_cap =Cargolanetype_average_capacity[5]
           
            if type_cap != 0:
                Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] /  Cargolanetype_average_capacity[5]))
            else:
                type_cap = Cargolanetype_average_capacity[4]
                if type_cap != 0:
                    Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] / Cargolanetype_average_capacity[4]))
                else:   
                    type_cap = Cargolanetype_average_capacity[3]
                    if type_cap != 0:
                        Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] /  Cargolanetype_average_capacity[3]))
                    else:
                        Product_max_cargolanenum.append(0)
                        
        elif product_typenum[i] == 4:
            type_cap = Cargolanetype_average_capacity[5]
            if type_cap != 0:
                Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] /  Cargolanetype_average_capacity[5]))
            else:
                type_cap = Cargolanetype_average_capacity[4]
                if type_cap != 0:
                    Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] /Cargolanetype_average_capacity[4]))
                else:   
                    Product_max_cargolanenum.append(0)
           
        elif product_typenum[i] == 5:
            type_cap = Cargolanetype_average_capacity[5]
            if type_cap != 0:
                Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] / Cargolanetype_average_capacity[5]))
            else:
                Product_max_cargolanenum.append(0)
            #Product_max_cargolanenum.append(math.ceil(replenishment_per_time[i] / Cargolanetype_average_capacity[product_typenum[i]]))
    # print("Product Max=", Product_max_cargolanenum)  
    
    return ID_CargoLane1, ID_CargoLane2, ID_CargoLane3, ID_CargoLane4, ID_CargoLane5, \
           Price_CargoLane1, Price_CargoLane2, Price_CargoLane3, Price_CargoLane4, Price_CargoLane5, \
           Sales_CargoLane1, Sales_CargoLane2, Sales_CargoLane3, Sales_CargoLane4, Sales_CargoLane5, \
           Product_max_cargolanenum, product_typenum, cargolane_should_empty,\
           New_ID1, New_ID2, New_ID3, New_ID4, New_ID5, Brand_CargoLane1, Brand_CargoLane2, Brand_CargoLane3, Brand_CargoLane4, Brand_CargoLane5,\
           setup_cost, replenishment_cost
   
#%%
Final_TS_revenue=[]
Final_revenue_of_no_repeated_product=[]
Final_Best_Solution=[]
Final_Best_Revenue=[]
Final_Best_revenue_of_no_repeated_product=[]
TS_list_revenue=[]
No_repeated_comparison=[]
revenue_comparison=[]


def Main_Program(Demand_Product_ID, Product_ID, Product_Volume, Product_Type, Product_Price, Demand_Product_Sales, product_typenum, Product_max_cargolanenum, CargoLane_ID, CargoLane_Type, setup_cost, replenishment_cost, total_cost, Product_Cost):
    
    ##---------------------- Initial Selection - Solution ----------------------------##
    
        
        copy_df_Product_info=df_Product_info.copy()
        copy_df_Product_demand=df_Product_demand.copy()
        
        copy_Product_ID= Product_ID.copy()
        copy_Product_Price= Product_Price.copy()
        copy_Product_Type= Product_Type.copy()
        copy_Product_Volume= Product_Volume.copy()
        
        copy_Demand_Product_ID= Demand_Product_ID.copy()
        copy_Demand_Product_Sales= Demand_Product_Sales.copy()
        
        demand_product_price=[]
        demand_product_type=[]
        demand_product_volume=[]
        demand_product_profit=[]
        demand_product_profit_product_max=[]
        demand_product_total_cost=[]
        demand_product_setup_cost=[]
        demand_product_replenishment_cost=[]
        demand_product_cost=[]
        
        Final_Selected_Product=[]
        Selected_Product=[]
        Selected_Product_Revenue=[]
        Cargo_Lane_Id_empty=[]
        In_Sol_Type=[]
        In_Sol_Volume=[]
        
        # print("Total_Cost=",total_cost)
        # print("setup cost=",setup_cost)
        #print("replenishment cost=",replenishment_cost)
        #print("Product cost=", Product_Cost)
       
        
        #add Product Price, Product Type, Product Volume data based on Demand_Product_ID
        for i in range (len(copy_Demand_Product_ID)):
            if i<=len(copy_Demand_Product_ID):
                for j in range(len(copy_Product_ID)):
                    if j<=len(copy_Product_ID):
                        if copy_Demand_Product_ID[i]==copy_Product_ID[j]:
                            demand_product_price.append(copy_Product_Price[j])
                            demand_product_type.append(copy_Product_Type[j])
                            demand_product_volume.append(copy_Product_Volume[j])
                            demand_product_total_cost.append(total_cost[j])
                            demand_product_setup_cost.append(setup_cost[j])
                            demand_product_replenishment_cost.append(replenishment_cost[j])
                            demand_product_cost.append(Product_Cost[j])
            else:
                break
            
        # print("tes",demand_product_total_cost)
        
        #add profit data
        # print("a=", copy_Demand_Product_Sales)
        # print("b=",demand_product_price)
        # print("c=",Product_Cost)
        # print("d=", demand_product_total_cost)
        
       
        for num1,num2, num3,num4, num5, num6 in zip(demand_product_price, demand_product_cost, copy_Demand_Product_Sales, demand_product_total_cost, demand_product_setup_cost, demand_product_replenishment_cost):
            demand_product_profit.append(((num1-num2)*num3)-num4-num5-num6)
        
        for num1,num2 in zip(demand_product_profit, Product_max_cargolanenum):
            demand_product_profit_product_max.append(num1*num2)
        
        
        #add data to df_Product_demand
        copy_df_Product_demand['Price']= demand_product_price
        copy_df_Product_demand['Type']= demand_product_type
        copy_df_Product_demand['Volume']= demand_product_volume
        copy_df_Product_demand['Product_Cost']= demand_product_cost
        copy_df_Product_demand['total_cost']= demand_product_total_cost
        
        copy_df_Product_demand['Revenue_Product']= demand_product_profit
        copy_df_Product_demand['Revenue_Product_Max']= demand_product_profit_product_max
        copy_df_Product_demand['Type_Number']= product_typenum
        copy_df_Product_demand['Max_Product_in_CL']= Product_max_cargolanenum
    
        
        #sort data based on Profit_Product_Max
        copy_df_Product_demand = copy_df_Product_demand.sort_values(by=['Revenue_Product_Max'],ascending=False) # using descending method
          
        #classify based on sorted data
        Demand_Product_ID = copy_df_Product_demand["Product_ID"].tolist()
        Demand_Product_Type = copy_df_Product_demand["Type"].tolist()
        Demand_Product_Volume = copy_df_Product_demand["Volume"].tolist()
        Demand_Product_Type_Num = copy_df_Product_demand["Type_Number"].tolist()
        Product_max_cargolanenum= copy_df_Product_demand["Max_Product_in_CL"].tolist()
        Demand_Product_Revenue = copy_df_Product_demand["Revenue_Product"].tolist()
        Demand_Product_Price= copy_df_Product_demand['Price'].tolist()
        Demand_Average_Sales= copy_df_Product_demand['Average_sales_month'].tolist()
        Demand_Product_Cost= copy_df_Product_demand['Product_Cost'].tolist()
        Demand_total_cost= copy_df_Product_demand['total_cost'].tolist()
        
        #print("DP_profit=", demand_product_profit)
        print("Demand_Product_Revenue=", Demand_total_cost)
        
        # Selection using Demand_Product_Type_Number
    
        for i in range(len(CargoLane_ID)):
            if i<=len(CargoLane_ID):
                for j in range(len(Demand_Product_ID)):
                    if j<=len(Demand_Product_ID):
                        if CargoLane_Type[i]== 0:
                            Initial_Chosen_Product= "None"
                            Final_Selected_Product.append(Initial_Chosen_Product)
                            Selected_Product_Revenue.append(int(0))
                            In_Sol_Type.append("None")
                            In_Sol_Volume.append(0)
                            break
                        
                        elif CargoLane_Type[i]== 1 and Demand_Product_Type_Num[j]==1:
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                                
                        elif CargoLane_Type[i]== 2 and Demand_Product_Type_Num[j]<=2:
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                            
                        elif CargoLane_Type[i]== 3 and Demand_Product_Type_Num[j]==3:
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                            
                        elif CargoLane_Type[i]== 4 and Demand_Product_Type_Num[j]>=3 and Demand_Product_Type_Num[j]<=4 :
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                                
                        
                        elif CargoLane_Type[i]== 5 and Demand_Product_Type_Num[j]>=3 and Demand_Product_Type_Num[j]<=5 :
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                                      
                                 
                else:
                     Initial_Chosen_Product='Empty'
                     Cargo_Lane_Id_empty.append(CargoLane_ID[i])
                     Final_Selected_Product.append(Initial_Chosen_Product)
                     Selected_Product_Revenue.append(int(0))
                     In_Sol_Type.append("Empty")
                     In_Sol_Volume.append("Empty")
                    
            else:
                break
            i=i+1
       
        '''
        #Selection using Demand_Product_Type and Demand_Product_Volume
        for i in range(len(CargoLane_ID)):
            if i<=len(CargoLane_ID):
                for j in range(len(Demand_Product_ID)):
                    if j<=len(Demand_Product_ID):
                        if CargoLane_Type[i]== 0:
                            Initial_Chosen_Product= "None"
                            Final_Selected_Product.append(Initial_Chosen_Product)
                            Selected_Product_Revenue.append(int(0))
                            In_Sol_Type.append("None")
                            In_Sol_Volume.append("None")
                            break
                        
                        elif CargoLane_Type[i]== 1 and Demand_Product_Type[j]=="CAN" and Demand_Product_Volume[j]<=330:
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                                
                        elif (CargoLane_Type[i]== 2 and Demand_Product_Volume[j]<=330 and Demand_Product_Type[j]=="CAN") or  (CargoLane_Type[i]== 2 and Demand_Product_Volume[j]<=330 and Demand_Product_Type[j]=="SCAN") :
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                            
                        elif CargoLane_Type[i]== 3 and Demand_Product_Volume[j]<=500 and Demand_Product_Type[j]=="PET":
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                            
                        elif CargoLane_Type[i]== 4 and Demand_Product_Volume[j]<=600 and Demand_Product_Type[j]=="PET" :
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                        
                        elif CargoLane_Type[i]== 5 and Demand_Product_Volume[j]<=600 and Demand_Product_Type[j]=="PET":
                            Initial_Chosen_Product= Demand_Product_ID[j]
                            Selected_Product_Count= Selected_Product.count(Demand_Product_ID[j])
                            if Selected_Product_Count>=Product_max_cargolanenum[j]:
                                pass
                            elif Selected_Product_Count<Product_max_cargolanenum[j]:
                                Final_Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product.append(Initial_Chosen_Product)
                                Selected_Product_Revenue.append( Demand_Product_Revenue[j])
                                In_Sol_Type.append(Demand_Product_Type[j])
                                In_Sol_Volume.append(Demand_Product_Volume[j])
                                break
                                      
                                 
                else:
                     Initial_Chosen_Product='Empty'
                     Cargo_Lane_Id_empty.append(CargoLane_ID[i])
                     Final_Selected_Product.append(Initial_Chosen_Product)
                     Selected_Product_Revenue.append(int(0))
                     In_Sol_Type.append("Empty")
                     In_Sol_Volume.append("Empty")
                    
            else:
                break
            i=i+1
        
        
        '''
        #print("Final_Selected_Product=", Final_Selected_Product)
        #Constraints of minimum total of chosen products
        Avg_Sales_Product_for_empty_CL= ((sum(Demand_Average_Sales)/ len(Demand_Average_Sales))/5) #!!!!
        # print(Avg_Sales_Product_for_empty_CL)
        
        
        index_product_havenot_chosen=[]
        repeated_id_in_selected_product=[] 
        randCargo_selected=[]
        randProd_selected=[]
        
        Minimum_Product_Chosen=6 #!!!
        Total_Product_Chosen= len(set(Selected_Product))
        #print("")
        # print("Total_Product_Chosen=",Selected_Product)
       
        if Total_Product_Chosen>= Minimum_Product_Chosen:
            Final_Selected_Product= Final_Selected_Product
            
        elif Total_Product_Chosen<Minimum_Product_Chosen:
            less_amount= Minimum_Product_Chosen - Total_Product_Chosen
            
            #define demand_product_id that has not been chosen yet
            for i in range(len(Demand_Product_ID)):
                if Demand_Product_ID[i] in Selected_Product:
                    pass
                else:
                    index_product_havenot_chosen.append(i)
            
            while Minimum_Product_Chosen>Total_Product_Chosen:
    
                randCargo= random.choice(range(0, len(CargoLane_ID)-1))        
                
                randProd= random.choice(index_product_havenot_chosen)
                # print(index_product_havenot_chosen)
               
                if randCargo  in randCargo_selected or randProd  in randProd_selected:
                 pass
                elif randCargo not in randCargo_selected and randProd not in randProd_selected:
                    
                    #define repeated demand_product_id in Selected Product
                    repeated_id_in_selected_product.clear()
                    for a in Final_Selected_Product:
                         n = Final_Selected_Product.count(a) # checking the occurrence of elements
                         
                         # if the occurrence is more than one we add it to the output list
                         if n > 1:
                      
                             if repeated_id_in_selected_product.count(a) == 0:  # condition to check
                      
                                 repeated_id_in_selected_product.append(a)
                   
                    
                    if Final_Selected_Product[randCargo] not in repeated_id_in_selected_product:
                        pass
                    elif Final_Selected_Product[randCargo] in repeated_id_in_selected_product:
                        if CargoLane_Type[randCargo]== 1 and Demand_Product_Type_Num[randProd]==1:
                            Final_Selected_Product[randCargo]=Demand_Product_ID[randProd]
                            #Selected_Product_Revenue[randCargo]= Demand_Product_Revenue[randProd]
                            Selected_Product_Revenue[randCargo]=(((Demand_Product_Price[randProd]- Demand_Product_Cost[randProd])* Demand_Average_Sales[randProd])- Demand_total_cost[randProd] - demand_product_setup_cost[randProd]- demand_product_replenishment_cost[randProd])
                            In_Sol_Type[randCargo]= Demand_Product_Type[randProd]
                            In_Sol_Volume[randCargo]=Demand_Product_Volume[randProd]
                            randCargo_selected.append(randCargo)
                            randProd_selected.append(randProd)
                         
                            Total_Product_Chosen= Total_Product_Chosen+1
                            
                        
                        elif CargoLane_Type[randCargo]== 2 and Demand_Product_Type_Num[randProd]<=2:
                            Final_Selected_Product[randCargo]=Demand_Product_ID[randProd]
                            #Selected_Product_Revenue[randCargo]= Demand_Product_Revenue[randProd]
                            Selected_Product_Revenue[randCargo]=(((Demand_Product_Price[randProd]- Demand_Product_Cost[randProd])* Demand_Average_Sales[randProd])- Demand_total_cost[randProd] - demand_product_setup_cost[randProd]- demand_product_replenishment_cost[randProd])
                            In_Sol_Volume[randCargo]=Demand_Product_Volume[randProd]
                            randCargo_selected.append(randCargo)
                            randProd_selected.append(randProd)
                           
                            Total_Product_Chosen= Total_Product_Chosen+1
                        
                        elif CargoLane_Type[randCargo]== 3 and Demand_Product_Type_Num[randProd]==3:
                            Final_Selected_Product[randCargo]=Demand_Product_ID[randProd]
                            #Selected_Product_Revenue[randCargo]= Demand_Product_Revenue[randProd]
                            Selected_Product_Revenue[randCargo]=(((Demand_Product_Price[randProd]- Demand_Product_Cost[randProd])* Demand_Average_Sales[randProd])- Demand_total_cost[randProd] - demand_product_setup_cost[randProd]- demand_product_replenishment_cost[randProd])
                            In_Sol_Volume[randCargo]=Demand_Product_Volume[randProd]
                            randCargo_selected.append(randCargo)
                            randProd_selected.append(randProd)
                            
                            Total_Product_Chosen= Total_Product_Chosen+1
                        
                        elif CargoLane_Type[randCargo]== 4 and Demand_Product_Type_Num[randProd]>=3 and Demand_Product_Type_Num[j]<=4:
                            Final_Selected_Product[randCargo]=Demand_Product_ID[randProd]
                            #Selected_Product_Revenue[randCargo]= Demand_Product_Revenue[randProd]
                            Selected_Product_Revenue[randCargo]=(((Demand_Product_Price[randProd]- Demand_Product_Cost[randProd])* Demand_Average_Sales[randProd])- Demand_total_cost[randProd] - demand_product_setup_cost[randProd]- demand_product_replenishment_cost[randProd])
                            In_Sol_Type[randCargo]= Demand_Product_Type[randProd]
                            In_Sol_Volume[randCargo]=Demand_Product_Volume[randProd]
                            randCargo_selected.append(randCargo)
                            randProd_selected.append(randProd)
                            
                            Total_Product_Chosen= Total_Product_Chosen+1
                            print(demand_product_total_cost[randProd])
                            
                        elif CargoLane_Type[randCargo]== 5 and Demand_Product_Type_Num[randProd]>=3 and Demand_Product_Type_Num[j]<=5 :
                            Final_Selected_Product[randCargo]=Demand_Product_ID[randProd]
                            #Selected_Product_Revenue[randCargo]= Demand_Product_Revenue[randProd]
                            Selected_Product_Revenue[randCargo]=(((Demand_Product_Price[randProd]- Demand_Product_Cost[randProd])* Demand_Average_Sales[randProd])- Demand_total_cost[randProd] - demand_product_setup_cost[randProd]- demand_product_replenishment_cost[randProd])
                            In_Sol_Type[randCargo]= Demand_Product_Type[randProd]
                            In_Sol_Volume[randCargo]=Demand_Product_Volume[randProd]
                            randCargo_selected.append(randCargo)
                            randProd_selected.append(randProd)
                           
                            Total_Product_Chosen= Total_Product_Chosen+1
                            
       
        #print(Selected_Product_Revenue)
        #print("Final_Selected_Product=", Final_Selected_Product)
        #print("")            
        #print("Cargo_Lane Empty=",Cargo_Lane_Id_empty)    
        
        #print("Total_Product_Chosen 2=",Total_Product_Chosen)
        #-- fill the empty cargolane--#
        
        #filter the df_product_info
        
        #print("for empty CL")
        Selected_id_for_empty_CG=[]
        Selected_id_for_empty_CG_for_1st_time=[]
        Selected_id_for_empty_CG_for_2nd_time=[]
        
        Avg_Sales_Product_for_empty_CL= ((sum(Demand_Average_Sales)/ len(Demand_Average_Sales))/5)
        
    
        #print(Demand_Product_Type_Num)
        #
        #print(CargoLane_Type)
        for i in range(len(Cargo_Lane_Id_empty)):
            if i<=len(Cargo_Lane_Id_empty):
                
                for j in range(len(Demand_Product_ID)):
                    if j<=len(Demand_Product_ID):
                        idx=Cargo_Lane_Id_empty[i]-1
                       
                        if CargoLane_Type[idx]== 1 and Demand_Product_Type_Num[j]==1:
                            initial_prod_for_empty_CG= Demand_Product_ID[j]
                            
                            #for checking the product to not be selected for more than twice
                        
                            if initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_2nd_time:
                                
                                pass
                            elif initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_1st_time:
                               
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_1st_time.append(initial_prod_for_empty_CG)
                                break
                            elif (initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_1st_time and initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_2nd_time):
                              
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_2nd_time.append(initial_prod_for_empty_CG)
                                break
                            
                        elif CargoLane_Type[idx]== 2 and Demand_Product_Type_Num[j]<=2:
                            initial_prod_for_empty_CG= Demand_Product_ID[j]
                            
                       
                            #for checking the product to not be selected for more than twice
                        
                            if initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_2nd_time:
                                
                                pass
                            elif initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_1st_time:
                             
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_1st_time.append(initial_prod_for_empty_CG)
                                break
                            elif (initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_1st_time and initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_2nd_time):
                               
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_2nd_time.append(initial_prod_for_empty_CG)
                                break
                            
                        elif CargoLane_Type[idx]== 3 and Demand_Product_Type_Num[j]==3:
                            initial_prod_for_empty_CG= Demand_Product_ID[j]
                            
                            #for checking the product to not be selected for more than twice
                        
                            if initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_2nd_time:
                               
                                pass
                            elif initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_1st_time:
                              
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_1st_time.append(initial_prod_for_empty_CG)
                                break
                            elif (initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_1st_time and initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_2nd_time):
                               
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_2nd_time.append(initial_prod_for_empty_CG)
                                break
                        
                        elif CargoLane_Type[idx]== 4 and Demand_Product_Type_Num[j]>=3 and Demand_Product_Type_Num[j]<=4:
                            initial_prod_for_empty_CG= Demand_Product_ID[j]
                            
                            #for checking the product to not be selected for more than twice
                        
                            if initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_2nd_time:
                               
                                pass
                            elif initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_1st_time:
                               
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_1st_time.append(initial_prod_for_empty_CG)
                                break
                            elif (initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_1st_time and initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_2nd_time):
                                
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_2nd_time.append(initial_prod_for_empty_CG)
                                break
                            
                        elif CargoLane_Type[idx]== 5 and Demand_Product_Type_Num[j]>=3 and Demand_Product_Type_Num[j]<=5:
                            initial_prod_for_empty_CG= Demand_Product_ID[j]
                            
                            #for checking the product to not be selected for more than twice
                        
                            if initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_2nd_time:
                              
                                pass
                            elif initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_1st_time:
                             
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_1st_time.append(initial_prod_for_empty_CG)
                                break
                            elif (initial_prod_for_empty_CG in Selected_id_for_empty_CG_for_1st_time and initial_prod_for_empty_CG not in Selected_id_for_empty_CG_for_2nd_time):
                                
                                Final_Selected_Product[idx]= initial_prod_for_empty_CG
                                Selected_Product.append(initial_prod_for_empty_CG)
                                Revenue_product_in_empty_CL= (((Demand_Product_Price[j]- Demand_Product_Cost[j])* Demand_Average_Sales[j])- Demand_total_cost[j] - demand_product_setup_cost[j]- demand_product_replenishment_cost[j])
                                Selected_Product_Revenue[idx]=Revenue_product_in_empty_CL
                                In_Sol_Type[idx]= Demand_Product_Type[j]
                                In_Sol_Volume[idx]=Demand_Product_Volume[j]
                                Selected_id_for_empty_CG.append(initial_prod_for_empty_CG)
                                Selected_id_for_empty_CG_for_2nd_time.append(initial_prod_for_empty_CG)
                                break
                            
                    else:
                        break
            else:
                break
            i=i+1
            
      
        #print("")
        #print('*'*100)
        #print("")
        #print("Cargo_Lane Empty=",Cargo_Lane_Id_empty)
        #print("")
        #print("Selected_id_for_empty_CG=", Selected_id_for_empty_CG)
        #print("") 
        #print("")
        #print("Final_Selected_Product=", Final_Selected_Product)
        #print("")
            
         
        Total_Revenue=sum(Selected_Product_Revenue)
        #print("Total Revenue=", Total_Revenue)
           
        Total_Revenue_of_no_repeated_product= sum(set(Selected_Product_Revenue))
        #print("Total_Revenue_of_no_repeated_product=", Total_Revenue_of_no_repeated_product)
            
       
           
        
        ##------------------------------------------- Tabu Search ----------------------------------------------------------------------## !!!!
        
              
        Initial_solution= Final_Selected_Product.copy()

        In_Sol_Revenue= Selected_Product_Revenue.copy()
        
        Tabu_list=[]
        temp_sol1=[]
        temp_sol2=[]
        Best_Solution=[]
        TS_Revenue=[]
        
        iter = 1
        Terminate = 0
        max_iter= 50
        run_max= 50
        
        
        for i in range(run_max):
            while iter <=max_iter:
             
                #print('\n\n### iter {}### '.format(iter))
            
                cand_move1= random.choice(range(0, len(Initial_solution)-1))
                cand_move2= random.choice(range(0, len(Initial_solution)-1))
                cand_move= (cand_move1, cand_move2)
           
                
                Initial_solution= Initial_solution
                In_Sol_Type = In_Sol_Type
                In_Sol_Volume = In_Sol_Volume
                
                if cand_move in Tabu_list:
                    #print("   best_move: {} => Tabu => Inadmissible".format(cand_move))
                    Current_solution= Initial_solution
                    Current_Sol_Type= In_Sol_Type
                    Current_Sol_Volume= In_Sol_Volume
                    Current_Sol_Revenue= In_Sol_Revenue
                    iter= iter+1
               
                    break
                    
                    
                elif cand_move not in Tabu_list:
                    if Initial_solution[cand_move1]=="None" or Initial_solution[cand_move2]=="None":
                        #print("   best_move: {} => None => Inadmissible".format(cand_move))
                        Current_solution= Initial_solution
                        Current_Sol_Type= In_Sol_Type
                        Current_Sol_Volume= In_Sol_Volume
                        Current_Sol_Revenue= In_Sol_Revenue
                        iter= iter+1
                     
                     
                        break
                    elif Initial_solution[cand_move1]!="None" or Initial_solution[cand_move2]!="None":
                        temp_sol1.clear()
                        temp_sol2.clear()
                        while True:
                            if CargoLane_Type[cand_move1]==1 and In_Sol_Type[cand_move2]=="CAN" and In_Sol_Volume[cand_move2]<=330:
                                temp_sol1.append(1)
                                break
                            elif (CargoLane_Type[cand_move1]==2 and In_Sol_Volume[cand_move2]<=330 and In_Sol_Type[cand_move2]=="CAN") or ( CargoLane_Type[cand_move1]==2 and In_Sol_Volume[cand_move2]<=330  and In_Sol_Type[cand_move2]=="SCAN") :
                                temp_sol1.append(1)
                                break
                            elif CargoLane_Type[cand_move1]==3 and In_Sol_Volume[cand_move2]<=500 and In_Sol_Type[cand_move2]=="PET":
                                temp_sol1.append(1)
                                break
                            elif CargoLane_Type[cand_move1]== 4 and In_Sol_Volume[cand_move2]<=600 and In_Sol_Type[cand_move2]=="PET" :
                                temp_sol1.append(1)
                                break
                            elif CargoLane_Type[cand_move1]== 5 and In_Sol_Volume[cand_move2]<=600 and In_Sol_Type[cand_move2]=="PET":
                                temp_sol1.append(1)
                                break
                            else:
                                temp_sol1.append(0)
                                break
                            
                        while True:
                            if CargoLane_Type[cand_move2]==1 and In_Sol_Type[cand_move1]=="CAN" and In_Sol_Volume[cand_move1]<=330:
                                temp_sol2.append(1)
                                break
                            elif (CargoLane_Type[cand_move2]==2 and In_Sol_Volume[cand_move1]<=330 and In_Sol_Type[cand_move1]=="CAN") or ( CargoLane_Type[cand_move2]==2 and In_Sol_Volume[cand_move1]<=330  and In_Sol_Type[cand_move1]=="SCAN") :
                                temp_sol2.append(1)
                                break
                            elif CargoLane_Type[cand_move2]==3 and In_Sol_Volume[cand_move1]<=500 and In_Sol_Type[cand_move1]=="PET":
                                temp_sol2.append(1)
                                break
                            elif CargoLane_Type[cand_move2]== 4 and In_Sol_Volume[cand_move1]<=600 and In_Sol_Type[cand_move1]=="PET" :
                                temp_sol2.append(1)
                                break
                            elif CargoLane_Type[cand_move2]== 5 and In_Sol_Volume[cand_move1]<=600 and In_Sol_Type[cand_move1]=="PET":
                                temp_sol2.append(1)
                                break
                            else:
                                temp_sol2.append(0)
                                break
                        
                      
                        while True:
                            if temp_sol1[0]==1 and temp_sol2[0]==1:
                                best_move= cand_move
                                Tabu_list.append(best_move)
                                Initial_solution[cand_move1], Initial_solution[cand_move2]=  Initial_solution[cand_move2], Initial_solution[cand_move1]
                                In_Sol_Type[cand_move1], In_Sol_Type[cand_move2]= In_Sol_Type[cand_move2], In_Sol_Type[cand_move1]
                                In_Sol_Volume[cand_move1], In_Sol_Volume[cand_move2]= In_Sol_Volume[cand_move2], In_Sol_Volume[cand_move1]
                                In_Sol_Revenue[cand_move1], In_Sol_Revenue[cand_move2]=In_Sol_Revenue[cand_move2], In_Sol_Revenue[cand_move1]
                                Current_solution= Initial_solution
                                Current_Sol_Type= In_Sol_Type
                                Current_Sol_Volume= In_Sol_Volume
                                Current_Sol_Revenue= In_Sol_Revenue
                                
                                #print("   best_move: {}  =>  Admissible ".format(cand_move))
                                #print("   Current_Solution: {} ".format(Current_solution))
                                iter= iter+1
                               
                                
                                break
                            else:
                                #print("   best_move: {} => None => Inadmissible".format(cand_move))
                                Current_solution= Initial_solution
                                Current_Sol_Type= In_Sol_Type
                                Current_Sol_Volume= In_Sol_Volume
                                Current_Sol_Revenue= In_Sol_Revenue
                                iter= iter+1
                             
                                
                                break
                    
                Initial_solution= Current_solution
                In_Sol_Type= Current_Sol_Type
                In_Sol_Volume = Current_Sol_Volume
                In_Sol_Revenue= Current_Sol_Revenue
        Best_Solution.append(Current_solution)
        TS_Revenue.append(Current_Sol_Revenue)
        Total_TS_Revenue= sum(Current_Sol_Revenue)
        Total_TS_Revenue_of_no_repeated_product= sum(set(Current_Sol_Revenue))
       
        
        if len(Final_Best_Revenue)==0:
            Final_Best_Solution.append(Best_Solution)
            Final_Best_Revenue.append(Total_TS_Revenue)
            Final_Best_revenue_of_no_repeated_product.append(Total_TS_Revenue_of_no_repeated_product)
            TS_list_revenue.append(TS_Revenue)
            
        elif len(Final_Best_Revenue)==1:

            if Total_TS_Revenue>= Final_Best_Revenue[0]:
               Final_Best_Solution.clear()
               Final_Best_Revenue.clear()
               Final_Best_revenue_of_no_repeated_product.clear() 
               TS_list_revenue.clear()
               
               Best_Solution= Best_Solution
               Final_Best_Solution.append(Best_Solution)
               
               Total_TS_Revenue= Total_TS_Revenue
               Final_Best_Revenue.append(Total_TS_Revenue)
               
               Total_TS_Revenue_of_no_repeated_product=Total_TS_Revenue_of_no_repeated_product
               Final_Best_revenue_of_no_repeated_product.append(Total_TS_Revenue_of_no_repeated_product)
               
               TS_Revenue=TS_Revenue
               TS_list_revenue.append(TS_Revenue)
               
            
        Final_TS_revenue.append(Total_TS_Revenue)
        Final_revenue_of_no_repeated_product.append(Total_TS_Revenue_of_no_repeated_product)
        
        length_data_Final_revenue= len(Final_revenue_of_no_repeated_product)
        index_data_Final_revenue= length_data_Final_revenue-1
        index_No_repeated_comparison= len(No_repeated_comparison)-1
        if length_data_Final_revenue==1:
            No_repeated_comparison.append(Final_revenue_of_no_repeated_product[0])
        elif length_data_Final_revenue>1:
            if Final_revenue_of_no_repeated_product[index_data_Final_revenue]> No_repeated_comparison[index_No_repeated_comparison]:
                No_repeated_comparison.append(Final_revenue_of_no_repeated_product[index_data_Final_revenue])
            elif Final_revenue_of_no_repeated_product[index_data_Final_revenue]<= No_repeated_comparison[index_No_repeated_comparison]:
                No_repeated_comparison.append(No_repeated_comparison[index_No_repeated_comparison])
        
        length_data_Final_revenue= len(Final_TS_revenue)
        index_data_Final_revenue= length_data_Final_revenue-1
        index_revenue_comparison= len(revenue_comparison)-1
        if length_data_Final_revenue==1:
            revenue_comparison.append(Final_TS_revenue[0])
        elif length_data_Final_revenue>1:
            if Final_TS_revenue[index_data_Final_revenue]> revenue_comparison[index_revenue_comparison]:
                revenue_comparison.append(Final_TS_revenue[index_data_Final_revenue])
            elif Final_TS_revenue[index_data_Final_revenue]<= revenue_comparison[index_revenue_comparison]:
                revenue_comparison.append(revenue_comparison[index_revenue_comparison])
            
        #print("aa=", Final_revenue_of_no_repeated_product)
        #print("bb=", No_repeated_comparison)
        #print("")
        #print('.'*50 , "Performed iterations: {}".format(iter-1), "Best found Solution: {}".format(Best_Solution), "TS_Revenue: {}".format(Total_TS_Revenue), "TS_List_Revenue: {}".format(TS_Revenue),sep="\n")   
        

        return demand_product_price, demand_product_type, demand_product_volume, demand_product_profit, demand_product_profit_product_max, Final_Selected_Product, Selected_Product_Revenue, \
               Cargo_Lane_Id_empty, Selected_id_for_empty_CG, Selected_id_for_empty_CG_for_1st_time, Selected_id_for_empty_CG_for_2nd_time, In_Sol_Type, In_Sol_Volume, \
               Total_Revenue, Total_Revenue_of_no_repeated_product, Tabu_list, temp_sol1, temp_sol2, Best_Solution, TS_Revenue, Total_TS_Revenue,Final_Best_Solution, Final_Best_Revenue, Final_Best_revenue_of_no_repeated_product,TS_list_revenue
  
#%%
def get_last_three_letters(a_list):
    return ''.join([x[-8:-6].lower() for x in a_list if len(x) > 2])

def Op_l(df_result, get_ol):
    ol= df_result.iloc[0,2]      
    return ol

# read the in/output path, parameters setting, error log
time_start = time.time() # start to count the time 開始計時
# parameters setting
mode = str(2)
new_prod_ratio = int(1) # 5%

# inputpath = os.path.normpath(sys.argv[1])
# outputpath = os.path.normpath(sys.argv[2])
# termination = int(os.path.normpath(sys.argv[3]))

# for heuristic
# if mode == str(1):
#     termination = 2
# elif mode == str(2):
#     termination = 2
# elif mode == str(3):
#     termination = 2

if mode == str(1):
    termination = 400
elif mode == str(2):
    termination = 100
elif mode == str(3):
    termination = 400

inputpath = r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\File thesis nsop" # test
inputpath1 = r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\zzzzzz"

if mode == str(1):
    outputpath = r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\outthesisTS" # test
elif mode == str(2):
    outputpath = r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\outts1"  # test
    outputpath_compare= r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\outts1"
else:
    outputpath =r"C:\Users\Admin\iCloudDrive\KULYEAH\lab\naskah\Thesis\outthesisTS"  # test

# for heuristic
# iter_mode1 = mode == str(1) and termination == 1
# iter_mode2 = mode == str(2) and termination == 1
# iter_mode3 = mode == str(3) and termination == 1
# iter_def = iter_mode1 == True or iter_mode2 == True or iter_mode3 == True

#iter_mode1 = mode == str(1) and 20 <= termination <= 600
#iter_mode2 = mode == str(2) and 200 <= termination <= 600
#iter_mode3 = mode == str(3) and 200 <= termination <= 600
#iter_def = iter_mode1 == True or iter_mode2 == True or iter_mode3 == True
iter_def = True

today_std = datetime.date.today()
today_std = str(today_std.year * 10000 + today_std.month * 100 + today_std.day)

today_std_time = time.localtime()
today_std_time = time.strftime('%H%M%S', today_std_time)

today_std_for_property = datetime.date.today()
today_std_for_property = int(today_std_for_property.year * 10000 + today_std_for_property.month * 100 + today_std_for_property.day)

#now we will Create and configure logger 
if mode == str(1):
    logging.basicConfig(filename="std_mode1_" + today_std + today_std_time + ".log", format='%(asctime)s %(message)s', filemode='w')
elif mode == str(2):
    logging.basicConfig(filename="std_mode2_" + today_std + today_std_time + ".log", format='%(asctime)s %(message)s', filemode='w')
elif mode == str(3):
    logging.basicConfig(filename="std_mode3_" + today_std + today_std_time + ".log", format='%(asctime)s %(message)s', filemode='w')

#Let us Create an object 
logger=logging.getLogger() 

#Now we are going to Set the threshold of logger to DEBUG 
logger.setLevel(logging.DEBUG)

ok = 0
okno = 0
okno_list = []

print("This program is a property of National Taiwan University of Science and Technology." + "\n")
logger.info("This program is a property of National Taiwan University of Science and Technology." + "\n")

op_l=[]

if os.path.exists(inputpath1)==True:
     inputfile_list1 = os.listdir(inputpath1)
     inputfile_list1.sort(reverse=False)
     
     for file in inputfile_list1:
        
         input_csv1 = os.path.join(inputpath1, file)
         input_sheet1 = "Sheet 1"
         
         df_result = pd.read_csv(input_csv1)
         get_ol= list(df_result.columns)
         
         ol= Op_l(df_result,get_ol)
         op_l.append(ol)
         
     
     op_l_int =[]
     for element in op_l:
         value_ol = float(element)
         op_l_int.append(value_ol)
     
     # print("Oppo loss=",op_l_int) 
     
if os.path.exists(inputpath) and os.path.exists(outputpath) and today_std_for_property <= 20231231 and iter_def == True:
    inputfile_list = os.listdir(inputpath)
    for file in inputfile_list:
        Final_TS_revenue.clear()
        Final_revenue_of_no_repeated_product.clear()
        Final_Best_Solution.clear()
        Final_Best_Revenue.clear()
        Final_Best_revenue_of_no_repeated_product.clear()
        TS_list_revenue.clear()
        No_repeated_comparison.clear()
        revenue_comparison.clear()
        
        try:
            
            print(file)
            logger.info(file)
            
            input_excel = os.path.join(inputpath, file)
            input_sheet_VM = "VM_info"
            input_sheet_ProEast = "Product_info_東區"
            input_sheet_ProNotEast = "Product_info_非東區"
            df_VM_info = pd.read_excel(input_excel, sheet_name = input_sheet_VM) # input VM_info sheet
            
            cargolane_num_should_be = (df_VM_info["CargoLane_TotalNumber"].squeeze()).tolist()
            
            df_VM_info = df_VM_info.append({"CargoLane_TotalNumber": int(0)}, ignore_index = True)
            CargoLane_Site_ID_for_log = int((df_VM_info.loc[0, ["Site_ID"]].squeeze()))
            '''
            if type(df_VM_info["CargoLane_ID"].squeeze().tolist()) == float:
                if math.isnan((df_VM_info["CargoLane_ID"].squeeze()).tolist()) == True:
                    print("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:VM_info is empty" + "\n")
                    logger.error("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:VM_info is empty" + "\n")
                    continue
            
            CargoLane_TotalNumber_first = int(df_VM_info.at[0, "CargoLane_TotalNumber"])
            if len(cargolane_num_should_be) != CargoLane_TotalNumber_first:
                print("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:The number of Cargolanes is not same as CargoLane_TotalNumber" + "\n")
                logger.error("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:The number of Cargolanes is not same as CargoLane_TotalNumber" + "\n")
                continue
            '''    
            locate_ID = list(OrderedDict.fromkeys((df_VM_info.loc[:, "Device_ID"].squeeze()).tolist())) # 重複值刪除，顯示所有點位
            del locate_ID[-1]
            Index_strart = int(0)
            Index_end = int(df_VM_info.at[0, "CargoLane_TotalNumber"]) - 1
            # Index_end = max((df_VM_info["CargoLane_ID"].squeeze()).tolist()) - 1 # 直接抓最後一個ID - 1
            # Index_end = len((df_VM_info["CargoLane_ID"].squeeze()).tolist()) - 2
            
            today = datetime.date.today()
            today = str(today.year * 10000 + today.month * 100 + today.day)
                    
            input_sheet_ProDemand = "Product_demand"
                # print(Index_strart, ":", Index_end)
            
            total_cost, df_VM_info, df_Product_info, df_Product_demand, df_replacement_matrix, VM_ID, CargoLane_Device_ID, CargoLane_Site_ID, CargoLane_TotalNumber, CargoLane_ID, CargoLane_Type, CargoLane_Height_Max, CargoLane_Height_Min, CargoLane_Diameter_Max_1, CargoLane_Diameter_Min_1, CargoLane_Area, CargoLane_Capacity, Current_Product, Max_Prod_Cnt, Min_Prod_Cnt, CargoLane_Allow_Special, CargoLane_Average_Replenishment, CargoLane_Category_Rate, CargoLane_Brand_Rate, Product_ID, Product_Price, Product_Cost, Product_Product_sales, Product_Type, Product_Volume, Product_Length, Product_Width, Product_Height, Product_New, Product_Brand, Product_Category, Product_Specialsize, Demand_Product_ID, Demand_Product_Sales, replacement_matrix, Demand_zero = read_data(df_VM_info, input_excel, input_sheet_ProEast, input_sheet_ProNotEast, input_sheet_ProDemand, Index_strart, Index_end)

            
            #ID_CargoLane1, ID_CargoLane2, ID_CargoLane3, ID_CargoLane4, ID_CargoLane5, Price_CargoLane1, Price_CargoLane2, Price_CargoLane3, Price_CargoLane4, Price_CargoLane5, Sales_CargoLane1, Sales_CargoLane2, Sales_CargoLane3, Sales_CargoLane4, Sales_CargoLane5, Product_max_cargolanenum, demand_product_typenum, cargolane_should_empty, cargolane_type_num, New_ID1, New_ID2, New_ID3, New_ID4, New_ID5, Brand_CargoLane1, Brand_CargoLane2, Brand_CargoLane3, Brand_CargoLane4, Brand_CargoLane5, product_product_typenum, replenishment_per_time, New_profit1, New_profit2, New_profit3, New_profit4, New_profit5, Cost_CargoLane1, Cost_CargoLane2, Cost_CargoLane3, Cost_CargoLane4, Cost_CargoLane5, sID_CargoLane1, sID_CargoLane2, sID_CargoLane3, sID_CargoLane4, sPrice_CargoLane1, sPrice_CargoLane2, sPrice_CargoLane3, sPrice_CargoLane4, sSales_CargoLane1, sSales_CargoLane2, sSales_CargoLane3, sSales_CargoLane4, sCost_CargoLane1, sCost_CargoLane2, sCost_CargoLane3, sCost_CargoLane4, sNew_ID1, sNew_ID2, sNew_ID3, sNew_ID4, sNew_profit1, sNew_profit2, sNew_profit3, sNew_profit4, snID_CargoLane1, snID_CargoLane2, snID_CargoLane3, snID_CargoLane4, snPrice_CargoLane1, snPrice_CargoLane2, snPrice_CargoLane3, snPrice_CargoLane4, snSales_CargoLane1, snSales_CargoLane2, snSales_CargoLane3, snSales_CargoLane4, snCost_CargoLane1, snCost_CargoLane2, snCost_CargoLane3, snCost_CargoLane4, snNew_ID1, snNew_ID2, snNew_ID3, snNew_ID4, snNew_profit1, snNew_profit2, snNew_profit3, snNew_profit4, setup_cost, replenishment_cost, total_cost, total_Cost_CargoLane1, total_Cost_CargoLane2, total_Cost_CargoLane3, total_Cost_CargoLane4, total_Cost_CargoLane5,  total_sCost_CargoLane1, total_sCost_CargoLane2, total_sCost_CargoLane3, total_sCost_CargoLane4,  total_snCost_CargoLane1, total_snCost_CargoLane2, total_snCost_CargoLane3, total_snCost_CargoLane4 = classify_demand_product(Product_ID, Product_Type, Product_Volume, Product_Price, Demand_Product_ID, Demand_Product_Sales, CargoLane_Average_Replenishment, Product_New, Product_Brand, Product_Specialsize, Product_Cost)
            #Recommend_ID1, Recommend_ID2, Recommend_ID3, Recommend_ID4, Recommend_ID5, Recommend_price1, Recommend_price2, Recommend_price3, Recommend_price4, Recommend_price5, Recommend_cost1, Recommend_cost2, Recommend_cost3, Recommend_cost4, Recommend_cost5, sRecommend_ID1, sRecommend_ID2, sRecommend_ID3, sRecommend_ID4, sRecommend_price1, sRecommend_price2, sRecommend_price3, sRecommend_price4, sRecommend_cost1, sRecommend_cost2, sRecommend_cost3, sRecommend_cost4, snRecommend_ID1, snRecommend_ID2, snRecommend_ID3, snRecommend_ID4, snRecommend_price1, snRecommend_price2, snRecommend_price3, snRecommend_price4, snRecommend_cost1, snRecommend_cost2, snRecommend_cost3, snRecommend_cost4,total_Recommend_cost1, total_Recommend_cost2, total_Recommend_cost3, total_Recommend_cost4, total_Recommend_cost5, total_sRecommend_cost1, total_sRecommend_cost2, total_sRecommend_cost3, total_sRecommend_cost4,total_snRecommend_cost1, total_snRecommend_cost2, total_snRecommend_cost3, total_snRecommend_cost4 = classify_recommend_product(Product_ID, Product_Type, Product_Volume, Product_Price, Demand_Product_ID, Product_Cost, setup_cost, replenishment_cost, total_cost)
            ID_CargoLane1, ID_CargoLane2, ID_CargoLane3, ID_CargoLane4, ID_CargoLane5, Price_CargoLane1, Price_CargoLane2, Price_CargoLane3, Price_CargoLane4, Price_CargoLane5, Sales_CargoLane1, Sales_CargoLane2, Sales_CargoLane3, Sales_CargoLane4, Sales_CargoLane5, Product_max_cargolanenum, product_typenum, cargolane_should_empty, New_ID1, New_ID2, New_ID3, New_ID4, New_ID5, Brand_CargoLane1, Brand_CargoLane2, Brand_CargoLane3, Brand_CargoLane4, Brand_CargoLane5,setup_cost, replenishment_cost  = classify_demand_product(Product_ID, Product_Type, Product_Volume, Product_Price, Demand_Product_ID, Demand_Product_Sales, CargoLane_Average_Replenishment, Product_New, Product_Brand)
        
            print("Model is running...")
            logger.info("Model is running...")
            
            #output_final_result, output_final_summarization, cur_each_chro_profit, heu_each_chro_profit, output_heuristic_result, output_heuristic_summarization = main_program(ID_CargoLane1, ID_CargoLane2, ID_CargoLane3, ID_CargoLane4, ID_CargoLane5, Price_CargoLane1, Price_CargoLane2, Price_CargoLane3, Price_CargoLane4, Price_CargoLane5, Sales_CargoLane1, Sales_CargoLane2, Sales_CargoLane3, Sales_CargoLane4, Sales_CargoLane5, Product_max_cargolanenum, demand_product_typenum, cargolane_should_empty, New_ID1, New_ID2, New_ID3, New_ID4, New_ID5, Brand_CargoLane1, Brand_CargoLane2, Brand_CargoLane3, Brand_CargoLane4, Brand_CargoLane5, Recommend_ID1, Recommend_ID2, Recommend_ID3, Recommend_ID4, Recommend_ID5, Recommend_price1, Recommend_price2, Recommend_price3, Recommend_price4, Recommend_price5, new_prod_ratio, replacement_matrix, New_profit1, New_profit2, New_profit3, New_profit4, New_profit5, sID_CargoLane1, sID_CargoLane2, sID_CargoLane3, sID_CargoLane4, sPrice_CargoLane1, sPrice_CargoLane2, sPrice_CargoLane3, sPrice_CargoLane4, sSales_CargoLane1, sSales_CargoLane2, sSales_CargoLane3, sSales_CargoLane4, sCost_CargoLane1, sCost_CargoLane2, sCost_CargoLane3, sCost_CargoLane4, sNew_ID1, sNew_ID2, sNew_ID3, sNew_ID4, sNew_profit1, sNew_profit2, sNew_profit3, sNew_profit4, sRecommend_ID1, sRecommend_ID2, sRecommend_ID3, sRecommend_ID4, sRecommend_price1, sRecommend_price2, sRecommend_price3, sRecommend_price4, sRecommend_cost1, sRecommend_cost2, sRecommend_cost3, sRecommend_cost4, snID_CargoLane1, snID_CargoLane2, snID_CargoLane3, snID_CargoLane4, snPrice_CargoLane1, snPrice_CargoLane2, snPrice_CargoLane3, snPrice_CargoLane4, snSales_CargoLane1, snSales_CargoLane2, snSales_CargoLane3, snSales_CargoLane4, snCost_CargoLane1, snCost_CargoLane2, snCost_CargoLane3, snCost_CargoLane4, snNew_ID1, snNew_ID2, snNew_ID3, snNew_ID4, snNew_profit1, snNew_profit2, snNew_profit3, snNew_profit4, snRecommend_ID1, snRecommend_ID2, snRecommend_ID3, snRecommend_ID4, snRecommend_price1, snRecommend_price2, snRecommend_price3, snRecommend_price4, snRecommend_cost1, snRecommend_cost2, snRecommend_cost3, snRecommend_cost4, termination)
            demand_product_price, demand_product_type, demand_product_volume, demand_product_profit, demand_product_profit_product_max, Final_Selected_Product, Selected_Product_Revenue, Cargo_Lane_Id_empty,\
            Selected_id_for_empty_CG, Selected_id_for_empty_CG_for_1st_time, Selected_id_for_empty_CG_for_2nd_time, In_Sol_Type, In_Sol_Volume, Total_Revenue, Total_Revenue_of_no_repeated_product,\
            Tabu_list, temp_sol1, temp_sol2, Best_Solution, TS_Revenue, Total_TS_Revenue, Final_Best_Solution, Final_Best_Revenue, Final_Best_revenue_of_no_repeated_product, TS_list_revenue  = Main_Program(Demand_Product_ID, Product_ID, Product_Volume, Product_Type,\
            Product_Price, Demand_Product_Sales, product_typenum, Product_max_cargolanenum, CargoLane_ID, CargoLane_Type, setup_cost, replenishment_cost, total_cost, Product_Cost)
            
            run=100
            for x in range(run-1):
            
              Main_Program(Demand_Product_ID, Product_ID, Product_Volume, Product_Type, Product_Price, Demand_Product_Sales, product_typenum, Product_max_cargolanenum, CargoLane_ID, CargoLane_Type, setup_cost, replenishment_cost, total_cost, Product_Cost)
            
            print("")
            print('*#'*50)
            print("")
            #print("Final_Revenue_of_no_repeated_prodcut=", Final_revenue_of_no_repeated_product)
            #print("")
            #print("Final_TS_Revenue=", Final_TS_revenue)
            print("")
            print("Final_Best_Solution=", Final_Best_Solution)
            print("")
            print("Final_Best_Revenue=", Final_Best_Revenue)
            print("")
            print("Final_Best_revenue_of_no_repeated_product=",Final_Best_revenue_of_no_repeated_product)
            print("")
            
            '''
            print("Final_Best_Revenue=", Final_Best_Revenue)
            print("")
            print("Final_Best_revenue_of_no_repeated_product=",Final_Best_revenue_of_no_repeated_product)
            print("")
            '''
            #print(revenue_comparison)
            #index_3 = iter_maxprofit_fitness.index(max(iter_maxprofit_fitness))
            #costlist = []
            #for i in range(len(iter_maxchro[index_3])):
                #if iter_maxchro[index_3][i] == "" or iter_maxchro[index_3][i] == "empty":
                    #costlist.append(0)
                #else:
                    #costlist.append(Product_Cost[Product_ID.index(iter_maxchro[index_3][i])])
            #compare_result = {"TS_Revenue": No_repeated_comparison}
            compare_result = {"TS_Revenue": revenue_comparison}
            output_compare_result = pd.DataFrame(compare_result)
            
            nama_file=[]
            get_number_file=[]
            get_number_file.clear()
            get_number_file.append(file)
            nama_file.append(get_last_three_letters(get_number_file))
            idxx=int(nama_file[0])
            print("File=", idxx)
            # print("tes", Final_Best_Revenue)
            
            # op_l_idx=[]
            # op_l_idx.append(op_l)
            # print(op_l_idx)
            fitness=[]
            for num1,num2 in zip(Final_Best_Revenue, op_l): #!!!!
                fitness.append(num1-num2) 
            
            print("Fitness=", fitness)
            
            
            final_result = {"VM ID": VM_ID, "Device ID": CargoLane_Device_ID, "Site_ID": CargoLane_Site_ID, "CargoLane ID": CargoLane_ID, "Product selection": Final_Best_Solution[0][0] , "Product profit": TS_list_revenue[0][0], "cargo_type": CargoLane_Type, "current prod": Current_Product}
            output_final_result = pd.DataFrame(final_result)
            output_final_summarization = pd.DataFrame()
            output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "Value": Final_Best_Revenue[0], "Value_type": "revenue"}, ignore_index=True)
            output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "Value": op_l[0], "Value_type": "opportunity_loss"}, ignore_index=True)
            output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "Value": fitness[0], "Value_type": "fitness"}, ignore_index=True)
            
            for i in range(len( Cargo_Lane_Id_empty )):
                output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "Value":  Cargo_Lane_Id_empty[i] , "Value_type": "empty"}, ignore_index=True)
            # for j in range(len(recommend_prod)):
            #     output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "value": recommend_prod[j], "Value_type": "recommend"}, ignore_index=True)
            for j in range(len(Selected_id_for_empty_CG)):
                output_final_summarization = output_final_summarization.append({"Site ID": CargoLane_Site_ID[0], "Device ID": CargoLane_Device_ID[0], "Value": Selected_id_for_empty_CG[j], "Value_type": "recommend"}, ignore_index=True)
            #print(output_final_summarization)
            #print('##')
            #print("compare=",output_compare_result)
            
            outputpath_s = os.path.join(outputpath, today + '_' + file + "_" + mode + "_result2.csv") # 設定路徑及檔名
            outputpath_r = os.path.join(outputpath, today + '_' + file + "_" + mode + "_result1.csv") # 設定路徑及檔名
            outputpath_c = os.path.join(outputpath_compare, today + '_' + file + "_" + mode + "TS.csv") # 設定路徑及檔名
            output_final_result.to_csv(outputpath_r, sep = ",", index = False, encoding = "utf-8")
            output_final_summarization.to_csv(outputpath_s, sep = ",", header = False, index = False, encoding = "utf-8")
            output_compare_result.to_csv(outputpath_c, sep = ",", index = False, encoding = "utf-8")
            
            # Index_strart = Index_end + 1
            CargoLane_Quantity = int(df_VM_info.at[Index_strart, "CargoLane_TotalNumber"])
            # Index_end = Index_strart + CargoLane_Quantity - 1
            print("result:" + str(CargoLane_Site_ID_for_log) + ":Execution succeed" + "\n")
            logger.info("result:" + str(CargoLane_Site_ID_for_log) + ":Execution succeed" + "\n")
            
            # if heu_each_chro_profit > 1:
            #     ok += 1
            # else:
            #     okno += 1
            #     okno_list.append(file[20:])
                
        except:                   # 如果 try 的內容發生錯誤，就執行 except 裡的內容
            print("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:Incorrect input data" + "\n")
            logger.error("result:" + str(CargoLane_Site_ID_for_log) + ":Execution failed:Incorrect input data" + "\n")
            pass                  # 略過
else:
    if today_std_for_property > 20221231:
        print("The deadline of exection was met, it's exceeded 20221231")
    elif iter_def == False:
        print("The number of iteration is out of range")
    elif os.path.exists(inputpath) == False and os.path.exists(outputpath) == False:
        print(inputpath, "and", outputpath, "do not exist")
    elif os.path.exists(inputpath) == False:
        print(inputpath, "do not exist")
    elif os.path.exists(outputpath) == False:
        print(outputpath, "do not exist")

time_end = time.time()    # 結束計時
time_total = time_end - time_start    # 執行所花時間
print('Spend:', time_total, '(s)')

# error log priority
# VM_info is empty v
# The number of Cargolanes is not same as CargoLane_TotalNumber v
# Some products in Product_demand are not in Product_info v
# Average Replenishment v
# Current_Product are not same as Demand_Product 
# Current_Product has conflict between product and product v
# CargoLanes are not sufficient v
# Incorrect input data, execution failed v


#%%
# to do list and need to check/revise
# 1. 啟發解還是有可能不符合品項數限制式
# 2. 最大 單位利潤*銷量/貨道需求!!!!! 只適用mode2推薦品項? 因為一般選品是用隨機, 新品及mode3推薦品沒有貨道可以計算 (O)

# 4. in/output path 輸入 (O)
# 5. 貨道尺寸選擇
# 6. 新品來源調整 (O)
# 7. 輸入輸出檔名調整site_id_device_id (O)

# 8. Heuristic加入一組max選解 (O)
# 9. current加入population (O)
# 10. 推薦品項利潤算法 (O)
# 11. SIZE版本: mutation要改candidate2 跟index
# 12. SIZE版本: chromosome要改cargoID 跟CargoLane_ID要改int
# 13. SIZE: chro_occupied要改
# 14. SIZE: objective要改


# 
# min_sku有問題 然後要改product_product_type(OK) & while終止(OK) >> 現在要讓擺不下的設一個error

# Natalia research
# 1. mutation: 一般選品納入

