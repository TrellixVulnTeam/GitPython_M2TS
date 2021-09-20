#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Sep 16 11:41:12 2021

@author: jayackley
"""
import pandas as pd

df = pd.DataFrame(data = {'Unique_ID': ['21-000001','21-000002','21-00003','21-00004'], 
                                            'Office': ['Brooklyn','Brooklyn','Bronx','Bronx'],
                                            'Modified Deliverable Tally':['Tier 1','Brief','Tier 1','Brief']})

print(df)

city_pivot = pd.pivot_table(df,values=['Unique_ID'], index = ['Office'], columns = ['Modified Deliverable Tally'],aggfunc=lambda x: len(x.unique()))
city_pivot.reset_index(inplace=True)


city_pivot.loc['Total',1:3] = city_pivot.sum(axis=0)

print(city_pivot)