#throughout this code we used the following abrevation: 
        #PV -> Popular Vote
        #EV -> Electoral Vote
        #DEM -> Democratic Party
        #REP -> Republican Party
import gurobipy as grb
import pandas as pd
from gurobipy import GRB


def readData(filename):
    xls = pd.ExcelFile(filename)
    df2000 = pd.read_excel(xls, '2000')
    df2004 = pd.read_excel(xls, '2004')
    df2008 = pd.read_excel(xls, '2008')
    df2012 = pd.read_excel(xls, '2012')
    df2016 = pd.read_excel(xls, '2016')
    df2020 = pd.read_excel(xls, '2020')

    return df2000 , df2004 , df2008 , df2012 ,df2016 ,df2020 

def formulation(df):
    DEMTotalEV = df.at[51, 'DEM EV']
    REPTotalEV = df.at[51, 'REP EV']
    
    REPRepresentative = df.at[0, 'Republican Candidate']
    DEMRepresentative = df.at[0, 'Democratic Candidate']
    
    winnerPartyName = 'DEM' if DEMTotalEV > REPTotalEV else 'REP'  
    runnerupPartyName = 'DEM' if DEMTotalEV < REPTotalEV else 'REP' 
    runnerupEV = DEMTotalEV if DEMTotalEV < REPTotalEV else REPTotalEV
    
    
    dfWinnerStates = df.loc[df[winnerPartyName + ' EV'] > 0]
    n = len(dfWinnerStates.index) - 1 #the last row does not include states information 
    
    
    #define a Gurobi model
    model = grb.Model()
    model.modelSense = GRB.MINIMIZE
    
    # defining dictionary changePopularVote: {the state that the president won all of the EV: PV needed to switch the result}
    changePopularVote = {i:
            dfWinnerStates.iloc[i]["PV needed to switch the result"]
            for i in range(n)
            }
    
    
    winnerStateList = []
    for i in range(n):
        winnerStateList.append(dfWinnerStates.iloc[i]['State'])
        
    #print(winnerStateList)
    #define the variable: x[i] gets value 1 if it is choosen to switch it
    x = []
    for i in range(n):
            x.append(model.addVar(vtype = grb.GRB.BINARY, lb = 0, ub = 1, obj = changePopularVote[i],
                                  name = "x_{%g}" % i ))

    model.update()    
    
    #define the constraint
    model.addConstr( grb.quicksum(x[i] * dfWinnerStates.iloc[i]["EV"] for i in range(n)) >= 270 - runnerupEV)
    
    #optimize the model
    model.optimize()
     
    #print solution
    dfResult = pd.DataFrame(columns = ['State', winnerPartyName + ' Theoretical PV',
                                       winnerPartyName + ' Theoretical EV',
                                       runnerupPartyName + ' Theoretical PV',
                                       runnerupPartyName + ' Theoretical EV'
                                       ])
    
    
    status = model.getAttr(grb.GRB.Attr.Status)
    if status == grb.GRB.OPTIMAL:
        for i in range(n):
            if x[i].getAttr(grb.GRB.Attr.X) == 1:
                dfResult = dfResult.append({'State': winnerStateList[i],
                                            winnerPartyName + ' Theoretical PV' : dfWinnerStates.iloc[i][winnerPartyName + " PV"] - changePopularVote[i] ,
                                            winnerPartyName + ' Theoretical EV' :  0 ,
                                            runnerupPartyName + ' Theoretical PV' : dfWinnerStates.iloc[i][runnerupPartyName + " PV"] + changePopularVote[i],
                                            runnerupPartyName + ' Theoretical EV' : dfWinnerStates.iloc[i]["EV"]
                                            },  
                                           ignore_index = True) 
    
    #adding total popular votes info
    dfResult = dfResult.append({'State': "Total PV needed to switch the result",
                                            winnerPartyName + ' Theoretical PV' : model.objVal,
                                            'Republican Candidate': REPRepresentative,
                                            'Democratic Candidate': DEMRepresentative
                                            },  
                                           ignore_index = True) 
    
    return dfResult

if __name__ == '__main__':
    
    
    ###########################
    # Read Data and Prepare Output Data Structure 
    ###########################
    
    input_file_name = "USelections.xlsx"
    output_file_name = "electionCloseness.xlsx"

    dfs = readData(input_file_name)
    
    writer = pd.ExcelWriter(output_file_name, engine='xlsxwriter')
    sheetName = 2000
    for df in dfs:
        #print(df)
        dfResult = formulation(df)
        dfResult.to_excel(writer, sheet_name = str(sheetName))
        sheetName += 4
    writer.save()
