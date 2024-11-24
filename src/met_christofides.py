
import networkx as nx
import numpy as np
import time
from scipy.spatial import distance_matrix
from itertools import combinations

def tsp(points):
    start_time = time.time()

    # Crear una matriz de distancias
    dist_matrix = distance_matrix(points, points)
    n = len(points)

    # 1. Encontrar el Árbol de Expansión Mínima (MST)
    G = nx.Graph()
    for i in range(n):
        for j in range(i + 1, n):
            G.add_edge(i, j, weight=dist_matrix[i][j])
    mst = nx.minimum_spanning_tree(G)

    # 2. Encontrar los vértices de grado impar
    odd_degree_nodes = [v for v, degree in mst.degree() if degree % 2 == 1]

    # 3. Emparejamiento mínimo perfecto de los vértices de grado impar
    subgraph_odd = G.subgraph(odd_degree_nodes)
    min_weight_matching = nx.algorithms.matching.min_weight_matching(subgraph_odd, True)

    # 4. Combinar el MST y el emparejamiento mínimo
    multi_graph = nx.MultiGraph(mst)
    multi_graph.add_edges_from(min_weight_matching)

    # 5. Encontrar el ciclo euleriano en el multigráfo
    eulerian_tour = list(nx.eulerian_circuit(multi_graph))

    # 6. Convertir el ciclo euleriano en un ciclo hamiltoniano (eliminar vértices repetidos)
    visited = set()
    hamiltonian_path = []
    for u, v in eulerian_tour:
        if u not in visited:
            visited.add(u)
            hamiltonian_path.append(u)
    hamiltonian_path.append(hamiltonian_path[0])  # Volver al inicio

    # Calcular la distancia total del ciclo hamiltoniano
    total_distance = sum(dist_matrix[hamiltonian_path[i]][hamiltonian_path[i + 1]] for i in range(len(hamiltonian_path) - 1))

    # Tiempo de ejecución
    exec_time = time.time() - start_time

    return hamiltonian_path, exec_time, total_distance
