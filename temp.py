# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from typing import Text
from dataclasses import astuple, dataclass
import numpy as np
import pandas as pd
from tqdm import tqdm
import pulp as plp

from datetime import datetime as dt

#%%

# Loading inventory database
df_inventory = pd.read_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/csv/inv_nov.csv')

#%%

print(len(df_inventory[df_inventory['Expire Date'] > df_inventory['Report Date']]['Item ID'].unique()))

#%%


df_inventory['Report Date'] = pd.to_datetime(df_inventory['Report Date'])
#df_inventory = df_inventory[(df_inventory['Expire Date'] <= dt(2030,12,1)) & (df_inventory['Expire Date'] > df_inventory['Report Date'])]
#df_inventory = df_inventory[df_inventory['Expire Date'] > df_inventory['Report Date']]

df_inventory['Expire Date'] = pd.to_datetime(df_inventory['Expire Date'], errors = 'coerce')

df_inventory['DtE'] = df_inventory['Expire Date'] - df_inventory['Report Date']

df_inventory['Expire Date'] = df_inventory['Expire Date'].fillna(dt(2199,12,1))
#df_inventory['Average Item Daily Use'] = df_inventory['Average Item Daily Use'].fillna(1)

df_inventory['DtE'] = df_inventory['DtE'].fillna(df_inventory['DtE'].max())
df_inventory['DtE'] = df_inventory['DtE'].dt.days
df_inventory['ItE'] = (df_inventory['BU Qty On Hand'] / df_inventory['Average Item Daily Use']) - df_inventory['DtE'].astype('int')

df_inventory['DOI'] = df_inventory['BU Qty On Hand'] \
                    / df_inventory['Average Item Daily Use']

df_inventory['DOI'] = df_inventory['DOI'].fillna(0)
df_inventory['DOI'] = df_inventory['DOI'].replace(np.inf, 0)
df_inventory['ItE'] = df_inventory['ItE'].replace(np.inf, 0)
df_inventory['ItE'] = df_inventory['ItE'].fillna(0)
df_inventory.loc[(df_inventory.ItE < 0),'ItE'] = 0

df_inventory['Delta DOI'] = df_inventory['DOI'] - df_inventory['DOI Target']

# DOI Balance = Delta DOI times the average consumption rate
df_inventory['DOI Balance'] = df_inventory['Delta DOI'] * \
                              df_inventory['Average Item Daily Use']

df_inventory['Min Shipment $ value'] = 50

#%%

@dataclass
class Columns:
    """
    Specify inventory dataframe column names.

    Place for specifying column names used on Inventory Dataframe

    Attributes
    ----------
    bu_num : str
        Column name for Business Unit ID.
    sku_id : str
        Column name for SKU ID.
    doi_balance : str
        Column name for Items balance.
    can_transfer : str
        Column name for can transfer flag.
    can_receive : str
        Column name for can receive flag.
    """
    bu_num: Text = 'Inv BU'
    sku_id: Text = 'Item ID'
    doi_balance: Text = 'DOI Balance'
    can_transfer: Text = 'Can transfer inventory?'
    can_receive: Text = 'Can receive inventory/'
    min_ship:Text = 'Min Shipment $ value'
    price: Text = 'Price'
    ite: Text = 'ItE'
    dte: Text = 'DtE'
    avg: Text = 'Average Item Daily Use'
    doi_target: Text = 'DOI Target'
    bu_qty: Text = 'BU Qty On Hand'

def sol_matrix(df_inventory, item_id, config: Columns = Columns()):
    """
    Create the solution matrix.

    Based on the item ID of our inventory database, it
    creates the matrix that contains in the first line and column
    all the BU Numbers that exist, at the second line the maximum number
    of items that these BU's can receive, and at the second column the number
    of items that will expire. The rest if the lines will be used later on to store
    values to be transfered between BU's.

    OBSERVATION: The method assumes that the Inventory dataframe already has
    the values for items to send and receive.

    Parameters
    ----------
    df_inventory : pandas.core.frame.DataFrame
        Dataframe with inventory.
    item_id : int
        The identification number of the SKU that is being analyzed.

    Returns
    -------
    sol_space : numpy.matrix
        Numpy matrix to be used in the optimization model.

    """
    bu_num, sku_id, doi_balance, can_transfer, can_receive, min_ship, price, ite, dte, avg, doi_target, bu_qty = astuple(config)

    bu_transfer = list(df_inventory[(df_inventory[sku_id] == item_id) &
                                    (df_inventory[can_transfer] == 'Yes')
                                    ][bu_num].unique())
    bu_receive = list(df_inventory[(df_inventory[sku_id] == item_id) &
                                   (df_inventory[can_receive] == 'Yes')
                                   ][bu_num].unique())

    # Filtering our inventory database for lines with the SKU that is being
    # analyzed
    sku_df = df_inventory[df_inventory[sku_id] == item_id]

    # Groupinh it by BU Number so that we have a consolidated position
    # by business unit
    sku = sku_df.groupby([bu_num], as_index=False).agg({doi_balance: 'mean',
                                                        min_ship: 'mean',
                                                        price: 'mean',
                                                        avg: 'mean',
                                                        ite: 'mean',
                                                        dte: 'mean',
                                                        doi_target: 'mean',
                                                        bu_qty: 'mean'
                                                        })

    sku[doi_balance] = round(sku[doi_balance],0)

    bu_prov = [x for x in bu_transfer if x in list(sku[((sku[doi_balance] > 0) | ((sku[doi_balance] == 0) & (sku[ite] != 0)))][bu_num].unique())]
    bu_rec = [x for x in bu_receive if x in list(sku[bu_num].unique())]

    bu_rec = [x for x in bu_receive if x in list(sku[((sku[doi_balance]<0) &
                                                       (-1*sku[doi_balance] > sku[min_ship]/sku[price])) | ((sku[doi_balance]==0) & (sku[ite] == 0))][bu_num].unique())]


    bu_prov = [x for x in bu_transfer if x in list(sku[
            (sku[doi_balance] > 0) |
            (
                (sku[doi_balance] == 0) &
                (sku[ite] != 0)
                )
            ][bu_num].unique())]

    bu_rec = [x for x in bu_receive if x in list(sku[
            ((sku[doi_balance] < 0) &
            (-1 * sku[doi_balance] >
                sku[min_ship] / sku[price]
                )) |
            ((sku[doi_balance] == 0) &
            (sku[ite] == 0) &
            (sku[avg] != 0)
            )][bu_num].unique())]

    # Adding two lines to store at our matrix
    # the BU Number (for identification), and the number
    # of items to be transfered.
    # Creating a matrix with zeros (it will be used as a base)
    sol_space = np.matrix(np.zeros((len(bu_prov) + 7, len(bu_rec) + 9)))

    # For every BU in our list of BU's...
    for idx, bu_id in enumerate(bu_prov):

        # Adding to our matrix first column the identification ID of all BU's

        for bu_idx, _ in sku.iterrows():
            if (sku[bu_num].iloc[bu_idx] == bu_id) \
                & ((sku[doi_balance].iloc[bu_idx] > 0) | ((sku[doi_balance].iloc[bu_idx] == 0) & (sku[ite].iloc[bu_idx] != 0))):
                    if bu_id in bu_transfer:
                        sol_space[idx+7, 0] = bu_prov[idx]
                        sol_space[idx+7, 1] = sku[doi_balance].iloc[bu_idx]
                        sol_space[idx+7, 2] = (sku[min_ship].iloc[bu_idx])
                        sol_space[idx+7, 3] = (sku[price].iloc[bu_idx])
                        sol_space[idx+7, 4] = sku[avg].iloc[bu_idx]
                        sol_space[idx+7, 5] = sku[ite].iloc[bu_idx]
                        sol_space[idx+7, 6] = sku[dte].iloc[bu_idx]
                        sol_space[idx+7, 7] = sku[doi_target].iloc[bu_idx]
                        sol_space[idx+7, 8] = sku[bu_qty].iloc[bu_idx]
 

    # For every BU in our list of BU's...
    for idx, bu_id in enumerate(bu_rec):

        # Adding to our matrix first column the identification ID of all BU's

        for bu_idx, _ in sku.iterrows():
            if (sku[bu_num].iloc[bu_idx] == bu_id) \
                & ((sku[doi_balance].iloc[bu_idx] < 0) | ((sku[doi_balance].iloc[bu_idx] == 0) & (sku[ite].iloc[bu_idx] == 0))):
                    if bu_id in bu_receive:
                        sol_space[0, idx+9] = bu_rec[idx]
                        sol_space[1, idx+9] = sku[doi_balance].iloc[bu_idx] * -1
                        sol_space[2, idx+9] = (sku[min_ship].iloc[bu_idx])
                        sol_space[3, idx+9] = (sku[price].iloc[bu_idx])
                        sol_space[4, idx+9] = sku[avg].iloc[bu_idx]
                        sol_space[5, idx+9] = sku[doi_target].iloc[bu_idx]
                        sol_space[6, idx+9] = sku[bu_qty].iloc[bu_idx]

    return sol_space

#%%

class Model():
    """
    Model for inventory optimization.

    Defines the optimization model variables, constraints, objective function and then solves it.

    Attributes
    ----------
    smatrix : numpy.matrix
        Numpy matrix to store the optimized variables.
    smatrix_df : pandas.core.frame.DataFrame
        Pandas dataframe to store the optimized variables.
    item_id : int
        SKU identification number.
    n : int
        Number of variables to be obtimized.
    set_j : range(1,n)
        Number of columns that the matrix has.
    a : list
        Constraint that limits the maximum number of items to be
        sent for a given BU to be the number of items to expire.
    u : list
        Upper bound for the variables of our optimization model. It is also
        used to limit the number of items the receiving BU's can accommodate.
    prob : lpProblem
        Model that needs to be optimized

    Methods
    -------
    variables()
        Define variable space. Every variable has an lower and upper bound.
    constraints(x_vars)
        Define model constraints.
    objective(x_vars)
        Define objective function.
    opt_matrix(opt_df)
        Generates matrix with optimization results.
    solve()
        Solve optimization problem.

    """

    def __init__(self, smatrix):
        """
        Arguments used at the optimization model.

        Parameters
        ----------
        smatrix : numpy.matrix
            Matrix of size (n x n) with data containing the amounts
            of items from a given SKU that can be sent and received from one BU to
            another.

        Returns
        -------
        None.

        """
        self.smatrix = smatrix
        self.smatrix_df = pd.DataFrame(smatrix)

        # Size of the matrix
        self.set_i = range(len(self.smatrix_df.index) - 7)
        self.set_j = range(len(self.smatrix_df.columns) - 9)

        #[MODEL CONSTRAINTS]

        # Constraint that limits the maximum number of items to be
        # sent for a given BU to be the number of items to expire.
        # ex.: If BU A has 10 items to expire, then we can transfer
        #      at most 10 items to other BU's (we can't transfer, say, 20 items)
        self.max_provide = list(self.smatrix_df.iloc[7:, 1])

        # Upper bound for the variables of our optimization model. It is also
        # used to limit the number of items the receiving BU's can accommodate.
        # ex.:
        #    - Case 1: If BU A can only receive 3 items of a certain category,
        #              then the maximum quantity of items that can be
        #              transfered to it is going to be 3 items
        #              (that is the upper bound).
        #    - Case 2: If BU A can only receive 3 items of a certain category,
        #              then the maximum quantity that all providing BU's
        #              can transfer to BU A is going to be 3.
        self.max_receive = list(self.smatrix_df.iloc[1, 9:])

        # Minimum shipment value $
        # Assuming that the minimum shipment value limitation is based on the
        # receiving BU and not on the transfering BU.
        self.min_ship = list(self.smatrix_df.iloc[2, 9:])

        # SKU Price
        # Price of one unit of that given SKU
        self.price = list(self.smatrix_df.iloc[3, 9:])

        self.max_expire = list(self.smatrix_df.iloc[7:, 5])


        self.days_expire = list(self.smatrix_df.iloc[7:, 6])
        self.avg_cons = list(self.smatrix_df.iloc[4,9:])
        
        self.avg_cons_prov = list(self.smatrix_df.iloc[7:,4])
        
        
        self.target_receive = list(self.smatrix_df.iloc[5,9:])
        self.target_provide = list(self.smatrix_df.iloc[7:, 7])
        
        self.inv_provide = list(self.smatrix_df.iloc[7:, 8])
        self.inv_receive = list(self.smatrix_df.iloc[6, 9:])

        # Big M gives us a way of adding logical constraints to the model.
        # In our case here, we use it to tell the model not to transfer items
        # from one BU to another, if their total value doesn't exceed the
        # minimum shipment value.
        self.BIG_M = max(sum(self.max_receive), sum(self.max_provide)) * 11

        for idx, val in enumerate(self.max_receive):

            if self.max_receive[idx] == 0:

                self.max_receive[idx] = (max(self.days_expire) / self.avg_cons[idx])*1.5

        # Defines the Model
        self.prob = plp.LpProblem("Inventory_Optimization")

    def variables(self):
        """
        Define optimizable variables (it is a matrix of size n x n).

        Returns
        -------
        Matrix with variables to optimize.

        """
        # if x is Integer
        x_vars = {
            (i, j): plp.LpVariable(cat='Integer',
                                   lowBound=0,
                                   upBound=min(self.max_provide[i],
                                               self.max_receive[j]),
                                   name="x_{}_{}".format(i, j))
            for i in self.set_i for j in self.set_j
            }

        # Used for filtering out x_vars if their total value don't exceed the
        # minimum threshold of the shipment value.
        y_vars = {
            (i, j): plp.LpVariable(cat='Binary',
                                   lowBound=0,
                                   name="y_{}_{}".format(i, j))
            for i in self.set_i for j in self.set_j
            }

        return x_vars, y_vars

    def constraints(self, x_vars, y_vars):
        """
        Add the constraints to our model.

        Right now, there are three constraints to the model:
            (1) Maximum number of items a given BU can send.
            (same as summing up all values of a given line)
            (2) Maximum number of items a given BU can receive.
            (same as summing up all valus of a given column)
            (3) Value of items to be sent from one BU to another needs to be
            bigger than the shippment value.
        NOTE: If you want to add new constraints to the model, add them here.

        Returns
        -------
        None.

        """
        arr = np.array(self.days_expire)
        ordered_idx = np.argsort(arr)

        # (2) SECOND CONSTRAINT
        # c = column (so for every column in our model)
        for col_idx in self.set_j:
         #   self.prob += plp.lpSum([x_vars[row_idx, col_idx]
          #                          for row_idx in self.set_i]
           #                        ) <= self.max_receive[col_idx]

            # (3) THIRD CONSTRAINT
            # r = row (so for every row in our model)
            for row_idx in self.set_i:

                self.prob += plp.lpSum(x_vars[row_idx, col_idx]  \
                                       - self.min_ship[col_idx] / self.price[col_idx]) \
                     >= - self.BIG_M * (1 - y_vars[row_idx, col_idx])
                self.prob += plp.lpSum(x_vars[row_idx, col_idx]) <= \
                    self.BIG_M * (y_vars[row_idx, col_idx])

        # (1) FIRST CONSTRAINT
        # r = row (so for every row in our model)
        for row_idx in self.set_i:
            self.prob += plp.lpSum([x_vars[row_idx, col_idx]
                                    for col_idx in self.set_j]
                                   ) <= max(self.max_provide[row_idx], self.max_expire[row_idx])

        for col_idx in self.set_j:
            tot = 0
            for row_idx in ordered_idx:
                tot += plp.lpSum(x_vars[row_idx, col_idx] * self.avg_cons[col_idx]**-1)
                self.prob += plp.lpSum(x_vars[row_idx, col_idx] * self.avg_cons[col_idx]**-1 + tot) <= self.days_expire[row_idx]

    def objective(self, x_vars, y_vars):
        """
        Add to model the objective function to be used.

        The one that is being used at the moment tries to minimize the amount
        of items to be expired. So:
            For a given SKU it tries to get the equation presented bellow
            closer to zero:
                Σ[Items to Expire]
                - Σ[Items to send (variable that is being optimized)] → 0
        Returns
        -------
        None.

        """
        
        #[min(self.max_receive[col_idx], max(self.max_provide[row_idx], self.max_expire[row_idx])) -
        objective = plp.lpSum(plp.lpSum([(self.target_provide[row_idx] * self.avg_cons_prov[row_idx] - 
                                          (self.inv_provide[row_idx] - plp.lpSum([x_vars[row_idx, col_idx] for col_idx in self.set_j])) 
                                ) for row_idx in self.set_i]) + 
                              plp.lpSum([(self.inv_receive[col_idx] + plp.lpSum([x_vars[row_idx, col_idx]
                                                                                 for row_idx in self.set_i]) - 
                                self.target_receive[col_idx] * self.avg_cons[col_idx])
                                for col_idx in self.set_j]))

        # for minimization
        self.prob.sense = plp.LpMaximize
        self.prob.setObjective(objective)

    def opt_matrix(self, opt_df):
        """
        Transform the result of the optimization model into a matrix form.

        Makes it easier to the end user to read.

        Parameters
        ----------
        opt_df : pandas.core.frame.DataFrame
            Dataframe with the results of the optimization model.

        Returns
        -------
        matrix_df : numpy.matrix
            Matrix with the recomendations created by the optimization model.

        """
        for idx, _ in opt_df.iterrows():

            i = opt_df["column_i"].iloc[idx] + 7
            j = opt_df["column_j"].iloc[idx] + 9

            val = opt_df["solution_value"].iloc[idx]

            self.smatrix[i, j] = val

        matrix_df = pd.DataFrame(self.smatrix)

        return matrix_df

    def solve(self, item_id):
        """
        Solve optimization problem.

        Main method of our Model class. It calls all the other methods
        of the class, creates the model, defines its contraints, adds the
        objective function and optimizes it.

        Returns
        -------
        opt_df : pandas.core.frame.DataFrame
            Dataframe with the results of the optimization model.
        matrix_df : numpy.matrix
            Matrix with the recomendations created by the optimization model.

        """
        # solving with CBC
        # Defining the variables space
        x_vars, y_vars = self.variables()
        # Adding the objetive function to the model
        self.objective(x_vars,y_vars)
        # Adding the constraints to the model
        self.constraints(x_vars, y_vars)
        # Solving the optimization problem given the model parameters
        # and definitions
        self.prob.solve()

        if plp.LpStatus[self.prob.status] != "Optimal":
            print("Status:", plp.LpStatus[self.prob.status])

        # Converting the result to a more readable format
        opt_df = pd.DataFrame.from_dict(x_vars,
                                        orient="index",
                                        columns=["variable_object"])
        opt_df.index = pd.MultiIndex.from_tuples(opt_df.index,
                                                 names=["column_i",
                                                        "column_j"])
        opt_df.reset_index(inplace=True)

        opt_df["solution_value"] = opt_df["variable_object"].apply(lambda item: item.varValue)
        opt_df["Item ID"] = item_id

        # Creating the matrix for better visualization of the output
        matrix_df = self.opt_matrix(opt_df)

        return opt_df, matrix_df

#%%

smatrix = sol_matrix(df_inventory, 183905)

#%%

opt_df, matrix_df = Model(smatrix).solve(183905)

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

                    opt_df['Receiver BU'].iloc[row] = opt_matrix.iloc[0, opt_df['column_j'].iloc[row] + 9]
                    opt_df['Provider BU'].iloc[row] = opt_matrix.iloc[opt_df['column_i'].iloc[row] + 7, 0]

                optimization_list.append(opt_df)
                optmization_matrix_list[str(sku)] = opt_matrix

#%%

result = pd.concat(optimization_list)
result['tot'] = result['solution_value'] * result['Value']
result['Value'] = round(result['Value'],2)
print(result['tot'].sum())

#%%
result['Value'] = round(result['Value'],2)
result['tot'] = round(result['tot'],2)

result.to_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/model/combined_model/Results/results_V6.csv')

#%%

rdf = pd.read_csv('/Users/erikingwersen/Desktop/EY-Quest-Diagnostics/model/combined_model/Results/results_V5.csv')

#%%


