import csv
import os
from surveyemaildata import agent_list
import pandas as pd
import numpy as np



# fileLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\MTD Detail - Jackson_18633_1504894770_2017-09-01_2017-09-08.csv'
fileLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\sales-funnel.xlsx'

df = pd.read_excel(fileLocation)
# print(df.head())

df["Status"] = df["Status"].astype("category")
df["Status"].cat.set_categories(["won","pending","presented","declined"],inplace=True)
print(pd.pivot_table(df,index=["Name"]))

# # print('current directory: ', os.getcwd())
# dataFile = open('C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\MTD Detail - Jackson_18633_1504894770_2017-09-01_2017-09-08.csv')
# fileReader = csv.reader(dataFile)
# # print(list(fileReader))
# # for row in fileReader:
# # 	print('Row #', str(fileReader.line_num), row[2], '\n')

# data = list(fileReader)
# data.pop(0)
# # print(data)

# agentData = []

# #After loop below, each agent Data item will have the format:
# #[AgentName, (Overall Experience (CSAT), Agent Satisfaction (ASAT), 
# #  Effort: Made Easy (CE)), Supervisor]
# for row in data:
# 	agentName = row[2].upper()
# 	CSAT = int(row[7])
# 	ASAT = int(row[13])
# 	effort = int(row[8])
# 	Supervisor = agent_list[agentName]

# 	agentData.append([agentName, CSAT, ASAT, effort, Supervisor])

# # print(len(data))


# agentData.sort()

# # for agent in agentData:
# # 	print(agent)

# overallExperience = []
