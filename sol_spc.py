# -*- coding: utf-8 -*-
"""
Created on Sun Nov 1 23:44:58 2020

@author: Erik Ingwersen
"""
from typing import Text
from dataclasses import astuple, dataclass
import numpy as np

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
    bu_num, sku_id, doi_balance, can_transfer, can_receive, min_ship, price = astuple(config)

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
                                                        price: 'mean'})

    sku[doi_balance] = round(sku[doi_balance],0)

    bu_prov = [x for x in bu_transfer if x in list(sku[(sku[doi_balance]>=0) &
                                                       (sku[doi_balance] > sku[min_ship]/sku[price])][bu_num].unique())]
    bu_rec = [x for x in bu_receive if x in list(sku[(sku[doi_balance]<0) &
                                                       (-1*sku[doi_balance] > sku[min_ship]/sku[price])][bu_num].unique())]

    # Adding two lines to store at our matrix
    # the BU Number (for identification), and the number
    # of items to be transfered.
    # Creating a matrix with zeros (it will be used as a base)
    sol_space = np.matrix(np.zeros((len(bu_prov) + 4, len(bu_rec) + 4)))

    # For every BU in our list of BU's...
    for idx, bu_id in enumerate(bu_prov):

        # Adding to our matrix first column the identification ID of all BU's

        for bu_idx, _ in sku.iterrows():
            if (sku[bu_num].iloc[bu_idx] == bu_id) & (
                sku[doi_balance].iloc[bu_idx] >= 0
            ) and bu_id in bu_transfer:
                sol_space[idx+4, 0] = bu_prov[idx]
                sol_space[idx+4, 1] = sku[doi_balance].iloc[bu_idx]
                sol_space[idx+4, 2] = (sku[min_ship].iloc[bu_idx])
                sol_space[idx+4, 3] = (sku[price].iloc[bu_idx])

    # For every BU in our list of BU's...
    for idx, bu_id in enumerate(bu_rec):

        # Adding to our matrix first column the identification ID of all BU's

        for bu_idx, _ in sku.iterrows():
            if (sku[bu_num].iloc[bu_idx] == bu_id) & (
                sku[doi_balance].iloc[bu_idx] < 0
            ) and bu_id in bu_receive:
                sol_space[0, idx+4] = bu_rec[idx]
                sol_space[1, idx+4] = sku[doi_balance].iloc[bu_idx] * -1
                sol_space[2, idx+4] = (sku[min_ship].iloc[bu_idx])
                sol_space[3, idx+4] = (sku[price].iloc[bu_idx])

    return sol_space
