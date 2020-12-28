import xlrd
import copy
import xlsxwriter
import os

selected_model = 1  # 1: max num_of_connections, budget per year; 2: max num_of_connections, budget per cluster; 3: min distance per cluster, budget per year

file = xlrd.open_workbook(os.getcwd()+"\\distances(time).xlsx")
sheet0 = file.sheet_by_index(0)
raw_data0 = [[sheet0.cell_value(r, c) for c in range(sheet0.ncols)] for r in range(sheet0.nrows)]
distances_time_data = copy.deepcopy(raw_data0)

file2 = xlrd.open_workbook(os.getcwd()+"\\fund.xlsx")
sheet1 = file2.sheet_by_index(0)
raw_data1 = [[sheet1.cell_value(r, c) for c in range(sheet1.ncols)] for r in range(sheet1.nrows)]
fund_data = copy.deepcopy(raw_data1)

file3 = xlrd.open_workbook(os.getcwd()+"\\links.xlsx")
sheet2 = file3.sheet_by_index(0)
raw_data2 = [[sheet2.cell_value(r, c) for c in range(sheet2.ncols)] for r in range(sheet2.nrows)]
links_data = copy.deepcopy(raw_data2)

container_port_name = []
port_cluster_name = []
port_clusters = []
port_cluster_index_of_port = []
year_name = []
budget_all_clusters_per_year = []
budget_each_cluster_per_year = []
total_investment_per_container_port = []
links_matrix = []
distances_time_matrix = []

budget_summary_starting_row = -1
stop_container_port_name_input = 0
for i in range(1,sheet1.nrows):
    if fund_data[i][0] == "":
        stop_container_port_name_input = 1
    if i>=1 and stop_container_port_name_input == 0:
        container_port_name.append(fund_data[i][1])
        if fund_data[i][2] not in port_cluster_name:
            port_cluster_name.append(fund_data[i][2])
        total_investment_per_container_port.append(fund_data[i][sheet1.ncols-1])
        port_cluster_index_of_port.append(port_cluster_name.index(fund_data[i][2]))
    if fund_data[i][0] == "" and budget_summary_starting_row == -1:
        budget_summary_starting_row = i
for i in range(len(port_cluster_name)):
    port_clusters_per_cluster = []
    for j in range(1,sheet1.nrows):
        if j>=1 and fund_data[j][0] != "":
            if fund_data[j][2] == port_cluster_name[i]:
                port_clusters_per_cluster.append(j-1)
    port_clusters.append(port_clusters_per_cluster)
for i in range(3,sheet1.ncols-1):
    year_name.append(str(int(fund_data[0][i])))
    budget_all_clusters_per_year.append(fund_data[len(fund_data)-4][i])
for i in range(budget_summary_starting_row,budget_summary_starting_row + len(port_cluster_name)):
    budget_each_cluster_all_years = []
    for j in range(3,sheet1.ncols-1):
        budget_each_cluster_all_years.append(fund_data[i][j])
    budget_each_cluster_per_year.append(budget_each_cluster_all_years)

for i in range(1,1+len(container_port_name)):
    links_matrix_per_row = []
    distances_time_matrix_per_row = []
    for j in range(2,2+len(container_port_name)):
        links_matrix_per_row.append(links_data[i][j])
        distances_time_matrix_per_row.append(distances_time_data[i][j])
    links_matrix.append(links_matrix_per_row)
    distances_time_matrix.append(distances_time_matrix_per_row)

import cplex
c=cplex.Cplex()

# Decision variable setting
num_of_variables = 0
for i in range(len(container_port_name)):
    for t in range(len(year_name)):
        c.variables.add(names= ['x'+str(container_port_name[i])+","+str(year_name[t])])
        c.variables.set_types([('x'+str(container_port_name[i])+","+str(year_name[t]), c.variables.type.continuous)])
        c.variables.set_lower_bounds([('x'+str(container_port_name[i])+","+str(year_name[t]), 0)])
        num_of_variables += 1

if selected_model == 1 or selected_model == 2:
    # Objective function (1)
    for i in range(len(container_port_name)):
        for t in range(len(year_name)):
            link_existence = 0
            for j in range(len(container_port_name)):
                if i < j:
                    link_existence += links_matrix[i][j]
                elif j < i:  # Do not consider cases with i = j
                    link_existence += links_matrix[j][i]
                if i == j:
                    continue
            c.objective.set_linear([('x'+str(container_port_name[i])+","+str(year_name[t]), link_existence * (len(year_name) - 1 - t))])
    c.objective.set_sense(c.objective.sense.maximize)

elif selected_model == 3:
    # Objective function (5)
    for i in range(len(container_port_name)):
        for t in range(len(year_name)):
            total_time = 0
            for g in range(len(port_clusters)):
                if i in port_clusters[g]:
                    for j in range(len(container_port_name)):
                        if i == j:
                            continue
                        if j in port_clusters[g]:
                            if i < j:
                                total_time += distances_time_matrix[i][j]
                            elif j < i:  # Do not consider cases with i = j
                                total_time += distances_time_matrix[j][i]
                    c.objective.set_linear([('x'+str(container_port_name[i])+","+str(year_name[t]), 
                                             (len(year_name) - 1 - t) / ((1/(len(port_clusters[g])-1) * total_time)))])
                    break
    c.objective.set_sense(c.objective.sense.maximize)

if selected_model == 1 or selected_model == 3:
    # Constraints (2)
    for t in range(len(year_name)):
        c.linear_constraints.add(names=['c'+str(year_name[t])])  # Set a constraint for each t
        c.linear_constraints.set_senses('c'+str(year_name[t]),"LE")  # Less than or equal sign
        c.linear_constraints.set_rhs('c'+str(year_name[t]),budget_all_clusters_per_year[t])  # Set the right hand side
        for a in range(num_of_variables):
            if year_name[t] in c.variables.get_names()[a]:  # Find variables that have the year index
                c.linear_constraints.set_coefficients('c'+str(year_name[t]),c.variables.get_names()[a],1)  # Sum the variables with their multipliers
            else:  # Do nothing for other variables
                pass

elif selected_model == 2:
    # Constraints (4)
    for g in range(len(port_clusters)):
        for t in range(len(year_name)):
            c.linear_constraints.add(names=['c'+str(port_cluster_name[g])+str(year_name[t])])  # Set a constraint for each t
            c.linear_constraints.set_senses('c'+str(port_cluster_name[g])+str(year_name[t]),"LE")  # Less than or equal sign
            c.linear_constraints.set_rhs('c'+str(port_cluster_name[g])+str(year_name[t]),budget_each_cluster_per_year[g][t])  # Set the right hand side
            for i in range(len(container_port_name)):
                for a in range(num_of_variables):
                    # Find variables that have the year index and container port in the selected cluster
                    if year_name[t] in c.variables.get_names()[a] and container_port_name[i] in c.variables.get_names()[a] and i in port_clusters[g]:
                        c.linear_constraints.set_coefficients('c'+str(port_cluster_name[g])+str(year_name[t]),c.variables.get_names()[a],1)  # Sum the variables with their multipliers
                    else:  # Do nothing for other variables
                        pass

# Constraints (3)
for i in range(len(container_port_name)):
    c.linear_constraints.add(names=['c'+str(container_port_name[i])])  # Set a constraint for each t
    c.linear_constraints.set_senses('c'+str(container_port_name[i]),"E")  # Equal sign
    c.linear_constraints.set_rhs('c'+str(container_port_name[i]),total_investment_per_container_port[i])  # Set the right hand side
    for a in range(num_of_variables):
        if container_port_name[i] in c.variables.get_names()[a]:  # Find variables that have the year index
            c.linear_constraints.set_coefficients('c'+str(container_port_name[i]),c.variables.get_names()[a],1)  # Sum the variables with their multipliers
        else:  # Do nothing for other variables
            pass

# LP Solve
c.write("formulation.lp")
c.parameters.timelimit = 60
start_time = c.get_time()
c.solve()
end_time = c.get_time()

# Write solutions to excel files
workbook = xlsxwriter.Workbook(os.getcwd()+'\\result_model'+str(selected_model)+'.xlsx')
worksheet = workbook.add_worksheet('result')

row = 0
col = 0

worksheet.write(row, 0, "time = ") 
worksheet.write(row, 1, end_time-start_time) 
worksheet.write(row+1, 0, "objective_value = ") 
worksheet.write(row+1, 1, c.solution.get_objective_value()) 

row = 3

# Printing titles of data
result_starting_row = row
for t in range(len(year_name)):
    worksheet.write(row, 0, "No.") 
    worksheet.write(row, 1, "Container Port") 
    worksheet.write(row, 2, "Cluster") 
    worksheet.write(row, col + 3 + t, year_name[t]) 
row += 1

investment_per_cluster_per_year = []
for i in range(len(port_cluster_name)):
    investment_per_cluster_per_year_a_row = []
    for t in range(len(year_name)):
        investment_per_cluster_per_year_a_row.append(0.0)
    investment_per_cluster_per_year.append(investment_per_cluster_per_year_a_row)
for i in range(len(container_port_name)):
    worksheet.write(row, 0, i+1) 
    worksheet.write(row, 1, container_port_name[i]) 
    worksheet.write(row, 2, port_cluster_name[port_cluster_index_of_port[i]]) 
    for t in range(len(year_name)):
        for a in range(num_of_variables):
            if container_port_name[i] in c.variables.get_names()[a] and year_name[t] in c.variables.get_names()[a]:
                worksheet.write(row, col + 3 + t, c.solution.get_values(a))
                investment_per_cluster_per_year[port_cluster_index_of_port[i]][t] += c.solution.get_values(a)
                break
    row += 1
for g in range(len(port_cluster_name)):
    worksheet.write(row, 1, "Total Cluster "+port_cluster_name[g]) 
    for t in range(len(year_name)):
        worksheet.write(row, 3+t, investment_per_cluster_per_year[g][t])
    row += 1

investment_effect_per_port_per_year = []
for i in range(len(container_port_name)):
    investment_effect_per_port_per_year_a_row = []
    for t in range(len(year_name)):
        investment_effect_per_port_per_year_a_row.append(0)
    investment_effect_per_port_per_year.append(investment_effect_per_port_per_year_a_row)

for i in range(len(container_port_name)):
    for t in range(len(year_name)):
        total_per_container_port_per_year = 0
        for a in range(num_of_variables):
            if container_port_name[i] in c.variables.get_names()[a] and year_name[t] in c.variables.get_names()[a]:
                for j in range(len(container_port_name)):
                    if i < j:
                        link_existence = links_matrix[i][j]
                    elif j < i:  # Do not consider cases with i = j
                        link_existence = links_matrix[j][i]
                    if i == j:
                        continue
                    total_per_container_port_per_year += link_existence * c.solution.get_values(a)
        for t2 in range(t+1, len(year_name)):
            investment_effect_per_port_per_year[i][t2] += total_per_container_port_per_year

row = result_starting_row
investment_effect_per_port_starting_column = 5+len(port_cluster_name)
worksheet.write(row-1, investment_effect_per_port_starting_column, "investment_effect_per_port") 
for t in range(len(year_name)):
    worksheet.write(row, investment_effect_per_port_starting_column + t, year_name[t]) 
row += 1
for i in range(len(container_port_name)):
    for t in range(len(year_name)):
        worksheet.write(row, investment_effect_per_port_starting_column+t, investment_effect_per_port_per_year[i][t])
    row += 1

row = result_starting_row
average_distance_with_ports_in_cluster_start_column = 5+len(port_cluster_name)*2+2
worksheet.write(row, average_distance_with_ports_in_cluster_start_column, "average_distance_with_ports_in_cluster") 
row += 1
for i in range(len(container_port_name)):
    total_time = 0
    for g in range(len(port_clusters)):
        if i in port_clusters[g]:
            for j in range(len(container_port_name)):
                if i == j:
                    continue
                if j in port_clusters[g]:
                    if i < j:
                        total_time += distances_time_matrix[i][j]
                    elif j < i:  # Do not consider cases with i = j
                        total_time += distances_time_matrix[j][i]
            worksheet.write(row, average_distance_with_ports_in_cluster_start_column, 1/(len(port_clusters[g])-1) * total_time)
            break
    row += 1

workbook.close()
