# Code originally written by Mohammad Javad Naderi (Oklahoma State University) in January 2021
# Revisions made by Austin Buchanan (Oklahoma State University) in February 2021

# Throughout this code, we use the following abrevations: 
#       EV      -> Electoral Votes
#       D / DEM -> Democratic Party
#       R / REP -> Republican Party
        
import gurobipy as gp
from gurobipy import GRB
import pandas as pd

def extract_data(df):
     
    # 50 states + DC
    n = 51    
    states = [df.iloc[i]["State"] for i in range(n)]
    
    # dictionary that maps 2-letter state codes to # electoral votes
    EV = {df.iloc[i]["State"] : df.iloc[i]["EV"] for i in range(n)}
    
    # dictionaries that map state codes to # votes for Republicans (R_votes) and Democrats (D_votes)
    D_votes = {df.iloc[i]["State"] : df.iloc[i]["DEM PV"] for i in range(n)}
    R_votes = {df.iloc[i]["State"] : df.iloc[i]["REP PV"] for i in range(n)}
    
    D_candidate = df.at[0, 'Democratic Candidate']
    R_candidate = df.at[0, 'Republican Candidate']  
    
    # number of electoral votes won by each candidate, assuming
    #   a plurarity suffices (not true in ME/NE) and no faithless electors
    D_EV = sum(EV[state] for state in states if D_votes[state]>R_votes[state])
    R_EV = sum(EV[state] for state in states if R_votes[state]>D_votes[state])
    
    return states, EV, D_votes, R_votes, D_candidate, R_candidate, D_EV, R_EV       


def print_data(states, EV, D_votes, R_votes, D_candidate, R_candidate, D_EV, R_EV ):
    
    print('{:<5s} {:>5s} {:>10s} {:>10s} {:^8s} {:^10s} {:>12s}'.format("State","EV","D_votes","R_votes","Winner","Runnerup","Votes-to-flip"))
    print("---------------------------------------------------------------------")
    for state in states:
        state_winner = "DEM" if D_votes[state]>R_votes[state] else "REP"
        state_runnerup = "DEM" if D_votes[state]<R_votes[state] else "REP"
        votes_to_flip = 1+abs(D_votes[state]-R_votes[state])
        print('{:<5s} {:>5d} {:>10d} {:>10d} {:^8s} {:^10s} {:>12d}'.format(state,EV[state],D_votes[state],R_votes[state],state_winner,state_runnerup,votes_to_flip))
    
    print("---------------------------------------------------------------------")

    overall_winner = D_candidate if D_EV > R_EV else R_candidate
    print("Democractic candidate",D_candidate,"won",D_EV,"electoral votes")
    print("Republican candidate",R_candidate,"won",R_EV,"electoral votes")
    print("---------------------------------------------------------------------")
    print(overall_winner,"won the election \n\n\n")
    

def solve_electoral_college_problem(df):
    
    states, EV, D_votes, R_votes, D_candidate, R_candidate, D_EV, R_EV = extract_data(df)
    print_data(states, EV, D_votes, R_votes, D_candidate, R_candidate, D_EV, R_EV)
    
    # who is the runnerup? write the model with respect to them
    if D_EV > R_EV: # D is winner; R is runnerup
        states_lost = [ state for state in states if D_votes[state]>R_votes[state] ]
        EV_won = R_EV
    else:           # R is winner; D is runnerup
        states_lost = [ state for state in states if R_votes[state]>D_votes[state] ]
        EV_won = D_EV
        
    votes_to_flip = {state: 1+abs(D_votes[state]-R_votes[state]) for state in states}
        
    # create a Gurobi model
    model = gp.Model()
    
    # create an X variable for each state that is lost, 
    #    where X[i]=1 means the runnerup "gets" the extra votes to flip state i
    X = model.addVars( states_lost, vtype=GRB.BINARY)
    
    # set the objective function (minimize number of extra votes needed)
    model.setObjective( gp.quicksum(votes_to_flip[state]*X[state] for state in states_lost), GRB.MINIMIZE)
    
    # set constraint that runnerup would need to reach 270 electoral college votes
    model.addConstr( gp.quicksum(EV[state]*X[state] for state in states_lost) >= 270 - EV_won)
        
    # #optimize the model
    model.optimize()
    
    # get solution (if solved properly)
    df_output = pd.DataFrame(columns = ['State', 'EV', 'Votes-to-flip'])
    if model.status == GRB.OPTIMAL:
        chosen_states = [state for state in states_lost if X[state].x>0.5 and votes_to_flip[state]>0]
        print("Chosen states:",chosen_states)
        
        for state in chosen_states:
            df_output = df_output.append({'State': state, 'EV':EV[state], 'Votes-to-flip': votes_to_flip[state]},ignore_index=True)
            
        EV_flipped = sum(EV[state] for state in chosen_states)
        PV_flipped = sum(votes_to_flip[state] for state in chosen_states)
        df_output = df_output.append({'State': 'Total flipped:', 'EV': EV_flipped, 'Votes-to-flip': PV_flipped},ignore_index=True)
        
    return df_output


if __name__ == '__main__':
        
    #################################################################################
    # Read data from Excel file, solve IP model, and then write output to Excel file
    #################################################################################
    
    input_filename = "election_data.xlsx"
    output_filename = "election_outputs.xlsx"
    
    xls = pd.ExcelFile(input_filename)
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    
    for election_year in range(2000,2024,4): # this is 2000, 2004, ...., 2020 in order
        
        df = pd.read_excel(xls,str(election_year))
        df_output = solve_electoral_college_problem(df)
        df_output.to_excel(writer, sheet_name = str(election_year), index = False)
        
    writer.save()
