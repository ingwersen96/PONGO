# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 13:55:00 2020

@author: Erik Ingwersen
"""

import pandas as pd
from tqdm import tqdm
from opt_model import Model
from sol_spc import sol_matrix

#%%

smatrix = sol_matrix(df_inventory, 100284)

print(len(smatrix))
#%%
opt_df, matrix_df = Model(smatrix).solve(100284)

#%%

# Loading inventory database
df_inventory = pd.read_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/csv/inv_nov.csv')

#%%

df_inventory['DOI'] = df_inventory['BU Qty On Hand'] \
                    / df_inventory['Average Item Daily Use']


df_inventory['DOI'] = df_inventory['DOI'].fillna(0)

# Calculating the delta DOI
df_inventory['Delta DOI'] = df_inventory['DOI'] - df_inventory['DOI Target']

# DOI Balance = Delta DOI times the average consumption rate
df_inventory['DOI Balance'] = df_inventory['Delta DOI'] * \
            df_inventory['Average Item Daily Use']

df_inventory['DOI Balance'] = df_inventory['DOI Balance'].fillna(0)

df_inventory.loc[((df_inventory['Average Item Daily Use'] == 0) &
                  (df_inventory['BU Qty On Hand'] > 0)), 'DOI Balance'] = df_inventory['BU Qty On Hand']

df_inventory['Min Shipment $ value'] = 50

#%%

df = df_inventory[((df_inventory['Average Item Daily Use'] == 0) &
                  (df_inventory['BU Qty On Hand'] > 0))]

#%%
# Running the model for entire inventory

# total number of unique SKU's
sku_list = list(df_inventory["Item ID"].unique())

optimization_list = []
optmization_matrix_list = {}


for sku in tqdm(sku_list):
    
    smatrix = sol_matrix(df_inventory, sku)
    
    # Filtering for SKU's that have at least one BU that can receive items
    # and one that can provide items.
    if smatrix[1,2:].sum() != 0: # Filtering first row
        if smatrix[2:,1].sum() != 0: # Filtering first column
            if smatrix[3,2:].sum() != 0: # Filtering price row
    
                opt_df, opt_matrix = Model(smatrix).solve(sku)
                opt_df["Value"] = df_inventory[df_inventory['Item ID'] == sku]['Price'].mean()
                opt_matrix[0,0] = sku
                
                opt_df = opt_df[opt_df['solution_value']>0]
                
                opt_df['Provider BU'] = 0
                opt_df['Receiver BU'] = 0
            
                for row in range(len(opt_df)):
                
                    opt_df['Receiver BU'].iloc[row] = opt_matrix.iloc[0, opt_df['column_j'].iloc[row] + 4]
                    opt_df['Provider BU'].iloc[row] = opt_matrix.iloc[opt_df['column_i'].iloc[row] + 4, 0]
                    
                optimization_list.append(opt_df)            
                optmization_matrix_list[str(sku)] = opt_matrix

#%%

result = pd.concat(optimization_list)
result['tot'] = result['solution_value'] * result['Value']
result['Value'] = round(result['Value'],2)
print(result['tot'].sum())

#%%

print(result[result['Item ID'] == 100012])

#%%

result.to_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/model/surplus_inv/Results/results_V5.csv')


#%%

rdf = pd.read_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/model/surplus_inv/Results/results_V3.csv')

#%%

rdf['Value'] = round(rdf['Value'],2)
rdf['tot'] = round(rdf['tot'],2)
#%%

rdf.to_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/model/surplus_inv/Results/results_V3.csv')

#%%


results_df = pd.read_csv('results_V2.csv')

#%%

print(results_df[results_df['tot']<50]['tot'].count())

#%%

df_test = df_inventory[df_inventory['Can transfer inventory?']=='No']

#%%

df_merged = df_test.merge(results_df, how='left', left_on=['Inv BU', 'Item ID'], right_on=['Provider BU', 'Item ID'])