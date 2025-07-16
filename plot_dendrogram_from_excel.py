
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, dendrogram

def plot_dendrogram_from_excel(excel_file_path, sheet_name=\'Stage 2 - Pairwise Similarity\', output_image_path=\'dendrogram_from_excel.png\'):
    """
    Построение дендрограммы из данных Excel.

    Args:
        excel_file_path (str): Путь к файлу Excel.
        sheet_name (str): Название листа в Excel, содержащего матрицу попарного сходства.
        output_image_path (str): Путь для сохранения изображения дендрограммы.
    """
    try:
        # Загрузка матрицы сходства из Excel
        similarity_matrix = pd.read_excel(excel_file_path, sheet_name=sheet_name, index_col=0)
    except FileNotFoundError:
        print(f"Ошибка: Файл Excel не найден по пути {excel_file_path}")
        return
    except KeyError:
        print(f"Ошибка: Лист \'{sheet_name}\' не найден в файле Excel.")
        return

    # Преобразование матрицы сходства в матрицу диссимилярности
    # Используем 7 - score, так как 7 - это максимальное сходство
    dissimilarity_matrix = 7 - similarity_matrix
    np.fill_diagonal(dissimilarity_matrix.values, 0) # Диссимилярность с самим собой равна 0

    # Проверка на симметричность и корректность данных
    if not np.allclose(dissimilarity_matrix, dissimilarity_matrix.T):
        print("Предупреждение: Матрица диссимилярности несимметрична. Будет усреднена.")
        dissimilarity_matrix = (dissimilarity_matrix + dissimilarity_matrix.T) / 2

    # Выполнение иерархической кластеризации
    # Используем метод 'ward' для минимизации дисперсии внутри кластеров
    linked = linkage(dissimilarity_matrix, method=\'ward\')

    # Построение дендрограммы
    plt.figure(figsize=(15, 10))
    dendrogram(
        linked,
        orientation=\'top\',
        labels=dissimilarity_matrix.index,
        distance_sort=\'descending\',
        show_leaf_counts=True
    )
    plt.title(\'Дендрограмма сходства городов\')
    plt.xlabel(\'Города\')
    plt.ylabel(\'Расстояние (диссимилярность)\')
    plt.tight_layout()
    plt.savefig(output_image_path)
    plt.close()

    print(f"Дендрограмма сохранена в {output_image_path}")

if __name__ == \'__main__\':
    # Пример использования:
    plot_dendrogram_from_excel(\'experimental_data.xlsx\', sheet_name=\'Stage 2 - Pairwise Similarity\', output_image_path=\'dendrogram_from_excel.png\')


