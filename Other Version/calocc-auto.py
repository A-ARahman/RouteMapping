import pandas as pd
import heapq
from openpyxl import load_workbook

# Read the Excel file
dataframe = pd.read_excel('datanode.xlsx', sheet_name='Data')

# Create a graph from the dataframe
graph = {}
for index, row in dataframe.iterrows():
    start, end, cost, occupancy = row['Node Awal'], row['Node Tujuan'], row['Cost'], row['Occupancy']
    if start not in graph:
        graph[start] = []
    if end not in graph:
        graph[end] = []
    graph[start].append((cost, end, 0, occupancy))  # Initialize bandwidth to 0
    graph[end].append((cost, start, 0, occupancy))  # If the graph is undirected

def dijkstra(graph, start):
    # Priority queue to store (cost, node)
    queue = [(0, start)]
    # Dictionary to store the shortest path to each node
    shortest_paths = {start: (None, 0, 0)}  # (previous node, cost, bandwidth)
    while queue:
        (cost, node) = heapq.heappop(queue)
        for edge_cost, neighbor, bandwidth, occupancy in graph[node]:
            old_cost = shortest_paths.get(neighbor, (None, float('inf'), 0))[1]
            new_cost = cost + edge_cost
            if new_cost < old_cost:
                shortest_paths[neighbor] = (node, new_cost, bandwidth)
                heapq.heappush(queue, (new_cost, neighbor))
    return shortest_paths

# Function to reconstruct the shortest path
def reconstruct_path(shortest_paths, start, end):
    path = []
    while end is not None:
        path.append(end)
        end = shortest_paths[end][0]
    path.reverse()
    return path

# Function to update bandwidth and occupancy
def update_bandwidth(graph, path, new_bandwidth, dataframe):
    updated_info = []
    for i in range(len(path) - 1):
        start, end = path[i], path[i + 1]
        for j in range(len(graph[start])):
            if graph[start][j][1] == end:
                cost, _, bandwidth, occupancy = graph[start][j]
                new_occupancy = occupancy + new_bandwidth
                graph[start][j] = (cost, end, bandwidth + new_bandwidth, new_occupancy)
                updated_info.append((start, end, bandwidth + new_bandwidth, new_occupancy))
                # Update the dataframe
                dataframe.loc[((dataframe['Node Awal'] == start) & (dataframe['Node Tujuan'] == end)) |
                              ((dataframe['Node Awal'] == end) & (dataframe['Node Tujuan'] == start)), 'Occupancy'] = new_occupancy
        for j in range(len(graph[end])):
            if graph[end][j][1] == start:
                cost, _, bandwidth, occupancy = graph[end][j]
                new_occupancy = occupancy + new_bandwidth
                graph[end][j] = (cost, start, bandwidth + new_bandwidth, new_occupancy)
    return updated_info

# Function to save results to the second sheet in the existing Excel file
def save_results_to_excel(file_path, dataframe, results):
    # Calculate the Idle column
    dataframe['Idle'] = dataframe['Bandwidth'] - dataframe['Occupancy']
    
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        dataframe.to_excel(writer, sheet_name='UpdatedData', index=False)
        results_df = pd.DataFrame(results, columns=['User ID', 'Path', 'Cost', 'Details'])
        results_df.to_excel(writer, sheet_name='Results2', index=False)

# Read input data from Excel file
input_data = pd.read_excel('datanode_input.xlsx', sheet_name='InputData')

results = []
for index, row in input_data.iterrows():
    user_id = row['ID']
    start_node = row['Node Awal']
    end_node = row['Node Tujuan']
    new_bandwidth = row['New Bandwidth']

    shortest_paths = dijkstra(graph, start_node)

    # Print shortest path and cost
    if end_node in shortest_paths:
        path = reconstruct_path(shortest_paths, start_node, end_node)
        cost = shortest_paths[end_node][1]
        updated_info = update_bandwidth(graph, path, new_bandwidth, dataframe)
        result = [user_id, ' -> '.join(path), cost]
        details = []
        for info in updated_info:
            start, end, bandwidth, occupancy = info
            details.append(f"{start}->{end}: Bandwidth = {bandwidth}, New Occupancy = {occupancy}")
        result.append('; '.join(details))
        results.append(result)
    else:
        results.append([user_id, None, None, "No path found"])

# Save the updated dataframe and results to the Excel file
save_results_to_excel('datanode_input.xlsx', dataframe, results)
