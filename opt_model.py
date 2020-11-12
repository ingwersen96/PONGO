# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 17:54:03 2020

@author: Erik Ingwersen
"""

import pandas as pd
import pulp as plp

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
        self.set_i = range(len(self.smatrix_df.index) - 4)
        self.set_j = range(len(self.smatrix_df.columns) - 4)

        #[MODEL CONSTRAINTS]

        # Constraint that limits the maximum number of items to be
        # sent for a given BU to be the number of items to expire.
        # ex.: If BU A has 10 items to expire, then we can transfer
        #      at most 10 items to other BU's (we can't transfer, say, 20 items)
        self.max_provide = list(self.smatrix_df.iloc[4:, 1])

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
        self.max_receive = list(self.smatrix_df.iloc[1, 4:])

        # Minimum shipment value $
        # Assuming that the minimum shipment value limitation is based on the
        # receiving BU and not on the transfering BU.
        self.min_ship = list(self.smatrix_df.iloc[2, 4:])

        # SKU Price
        # Price of one unit of that given SKU
        self.price = list(self.smatrix_df.iloc[3, 4:])
        
        
        # Big M gives us a way of adding logical constraints to the model.
        # In our case here, we use it to tell the model not to transfer items
        # from one BU to another, if their total value doesn't exceed the
        # minimum shipment value.
        self.BIG_M = max(sum(self.max_receive), sum(self.max_provide)) * 11
        
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
        # (2) SECOND CONSTRAINT
        # c = column (so for every column in our model)
        for col_idx in self.set_j:
            self.prob += plp.lpSum([x_vars[row_idx, col_idx]
                                    for row_idx in self.set_i]
                                   ) <= self.max_receive[col_idx]

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
                                   ) <= self.max_provide[row_idx]
    
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
        objective = plp.lpSum(min(sum(self.max_receive), sum(self.max_provide)) - plp.lpSum([x_vars[row_idx, col_idx]
                               for col_idx in self.set_j
                               for row_idx in self.set_i]))

        # for minimization
        self.prob.sense = plp.LpMinimize
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

            i = opt_df["column_i"].iloc[idx] + 4
            j = opt_df["column_j"].iloc[idx] + 4

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

        #print("Status:", plp.LpStatus[self.prob.status])

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
