
import pandas as pd
import numpy as np
from scipy.cluster.hierarchy import linkage, dendrogram
import matplotlib.pyplot as plt

def visualize_dendrogram(excel_file_path):
    df = pd.read_excel(excel_file_path)

    # Extract city names and subject scores
    cities = df["Город 1"].tolist() + df["Город 2"].tolist()
    cities = sorted(list(set(cities)))

    # Create a similarity matrix based on the average scores
    # Initialize with zeros
    similarity_matrix = pd.DataFrame(0.0, index=cities, columns=cities)

    for index, row in df.iterrows():
        city1 = row["Город 1"]
        city2 = row["Город 2"]
        # Calculate the average score for the pair
        avg_score = row.iloc[2:].mean() # Scores start from the 3rd column
        similarity_matrix.loc[city1, city2] = avg_score
        similarity_matrix.loc[city2, city1] = avg_score

    # Convert similarity to dissimilarity (distance)
    # A higher similarity score means a smaller distance
    # Assuming scores are 1-7, max_score = 7, min_score = 1
    # Distance = max_score - score
    max_score = 7
    distance_matrix = max_score - similarity_matrix

    # For self-similarity, distance is 0
    np.fill_diagonal(distance_matrix.values, 0)

    # Perform hierarchical clustering
    # Use 'average' linkage method
    linked = linkage(distance_matrix, method='average')

    # Plot the dendrogram
    plt.figure(figsize=(20, 10))
    dendrogram(
        linked,
        orientation='top',
        labels=cities,
        distance_sort='descending',
        show_leaf_counts=True
    )
    plt.title('Дендрограмма фонетического сходства городов')
    plt.xlabel('Города')
    plt.ylabel('Расстояние')
    plt.tight_layout()
    plt.savefig('dendrogram.png')
    plt.show()

if __name__ == '__main__':
    visualize_dendrogram('stage2_data.xlsx')


