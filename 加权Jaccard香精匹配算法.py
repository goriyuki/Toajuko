import pandas as pd
import pyodbc
import numpy as np
import matplotlib.pyplot as plt
from decimal import Decimal, getcontext
from collections import deque
import time

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

getcontext().prec = 4  # 可调精度

# 参数配置
excel_file = '你的精油配方.xlsx'
banned_file = '你的剔除列表.xlsx'

JACCARD_THRESHOLD = Decimal('0.5')
WEIGHT_STRATEGY = 'importance'
TOP_N = 30

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=192.168.1.10;"
    "DATABASE=Testapi;"
    "UID=sa;"
    "PWD=1227;"
)

def preprocess_uploaded_formula(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    df = df[['RMCode', 'ChemicalName', 'FID含量']]
    banned_df = pd.read_excel(banned_file, engine='openpyxl')
    banned_rmcodes = set(str(code).strip() for code in banned_df.iloc[:, 0] if pd.notnull(code))

    df['RMCode'] = df['RMCode'].astype(str).str.strip()
    mask = df['RMCode'].str.replace(' ', '', regex=False).str.lower() == 'normcode'
    df.loc[mask, 'RMCode'] = df.loc[mask, 'ChemicalName']

    df = df[~df['RMCode'].isin(banned_rmcodes)]
    df = df.groupby('RMCode', as_index=False).sum()
    df['FID含量'] = df['FID含量'].apply(Decimal)
    total = sum(df['FID含量'])
    df['FID_percent'] = df['FID含量'].apply(lambda x: (x / total) * Decimal(100))

    return {rm: val for rm, val in zip(df['RMCode'], df['FID_percent'])}

def load_database_formulas():
    star_timer = time.time()
    cursor = conn.cursor()
    cursor.execute(" SELECT b.RMCode, b.FID_percent,a.id,a.filename  FROM [00000_UploadedFormulas] a Left join [00000_FormulaDetails] b ON a.id = b.formula_id order by formula_id")
    # formula_meta = cursor.fetchall()
    # 获取列名
    columns = [column[0] for column in cursor.description]
    # 转换为字典列表
    formula_meta = [dict(zip(columns, row)) for row in cursor.fetchall()]
    formula_vectors = deque()
    filenames = deque()
    results = deque()
    index = None
    for row in formula_meta:
        if index is None:
            index = row['filename']
        if index != row['filename']:
            filenames.append(index)
            formula_dict = {row[0]: Decimal(row[1]) for row in results}
            formula_vectors.append(formula_dict)
            results = deque()
            index = row['filename']
        results.append((row['RMCode'],row['FID_percent']))
    #最后一组数据
    filenames.append(index)
    formula_dict = {row[0]: Decimal(row[1]) for row in results}
    formula_vectors.append(formula_dict)
    end_timer = time.time()
    print(f"读取数据库数据耗时：{end_timer - star_timer:.2f}秒")
    return formula_vectors, filenames

def proportion_weighted_jaccard(vec_a, vec_b):
    keys_a = set(vec_a.keys())
    keys_b = set(vec_b.keys())

    shared_keys = keys_a & keys_b
    only_a = keys_a - keys_b
    only_b = keys_b - keys_a

    numerator = Decimal(2) * sum(min(vec_a[k], vec_b[k]) for k in shared_keys)
    denominator = (
        sum(vec_a[k] for k in only_a) +
        sum(vec_b[k] for k in only_b) +
        Decimal(2) * sum(max(vec_a[k], vec_b[k]) for k in shared_keys)
    )

    return numerator / denominator if denominator > 0 else Decimal(0)

def weighted_euclidean_distance(vec_a, vec_b, weight_strategy='importance'):
    all_components = set(vec_a.keys()).union(set(vec_b.keys()))
    weighted_squared_sum = Decimal(0)

    for component in all_components:
        val_a = vec_a.get(component, Decimal(0))
        val_b = vec_b.get(component, Decimal(0))
        diff = abs(val_a - val_b)

        if weight_strategy == 'importance':
            weight = Decimal(1.0) + (max(val_a, val_b) / Decimal(10.0))
        elif weight_strategy == 'balanced':
            weight = Decimal(2.0) if (val_a > 0 and val_b > 0) else Decimal(1.0)
        else:
            weight = Decimal(1.0)

        weighted_squared_sum += (diff * weight) ** 2

    return weighted_squared_sum.sqrt()

def find_similar_formulas(uploaded_vector, database_vectors, filenames,
                          jaccard_threshold, weight_strategy, top_n):
    results = []

    for db_vector, filename in zip(database_vectors, filenames):
        jaccard_sim = proportion_weighted_jaccard(uploaded_vector, db_vector)
        if jaccard_sim < jaccard_threshold:
            continue

        weighted_dist = weighted_euclidean_distance(uploaded_vector, db_vector, weight_strategy)
        similarity_score = Decimal(1.0) / (Decimal(1.0) + weighted_dist)

        results.append({
            'filename': filename,
            'jaccard_similarity': jaccard_sim,
            'weighted_distance': weighted_dist,
            'similarity_score': similarity_score,
            'db_vector': db_vector
        })

    results.sort(key=lambda x: x['similarity_score'], reverse=True)
    return results[:top_n]

def plot_comparison(uploaded_vector, matched_vector, formula_name):
    all_rmcodes = sorted(set(uploaded_vector.keys()) | set(matched_vector.keys()))
    uploaded_vals = [float(uploaded_vector.get(rm, Decimal(0))) for rm in all_rmcodes]
    matched_vals = [float(matched_vector.get(rm, Decimal(0))) for rm in all_rmcodes]

    x = np.arange(len(all_rmcodes))
    width = 0.35

    fig, ax = plt.subplots(figsize=(max(12, len(all_rmcodes) * 0.3), 6))
    ax.bar(x - width / 2, uploaded_vals, width, label='上传配方', alpha=0.8)
    ax.bar(x + width / 2, matched_vals, width, label='相似配方', alpha=0.8)

    ax.set_xticks(x)
    ax.set_xticklabels(all_rmcodes, rotation=45, ha='right')
    ax.set_ylabel('FID_percent')
    ax.set_title(f'成分对比：上传配方 vs {formula_name}')
    ax.legend()
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()

def visualize_jaccard_difference(uploaded_vector, matched_vector, formula_name):
    uploaded_set = set(uploaded_vector.keys())
    matched_set = set(matched_vector.keys())

    only_uploaded = uploaded_set - matched_set
    only_matched = matched_set - uploaded_set
    shared = uploaded_set & matched_set

    categories = ['上传配方独有', '相似配方独有', '共同拥有']
    counts = [len(only_uploaded), len(only_matched), len(shared)]

    plt.figure(figsize=(6, 4))
    bars = plt.bar(categories, counts, color=['#ff9999', '#9999ff', '#99ff99'], alpha=0.8)

    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2.0, height + 0.2, f'{int(height)}', ha='center')

    plt.title(f"成分重合度分析：{formula_name}")
    plt.ylabel('成分数量')
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.show()

def print_results(results):
    print(f"\n{'排名':<4} {'配方名称':<30} {'加权杰卡德':<8} {'欧氏距离':<10} {'相似度':<8}")
    print("-" * 70)
    for i, result in enumerate(results, 1):
        print(f"{i:<4} {result['filename']:<30} "
              f"{result['jaccard_similarity']:.4f}   "
              f"{result['weighted_distance']:<10.4f} "
              f"{result['similarity_score']:.4f}")

# 主程序入口
if __name__ == "__main__":
    print("读取配方数据...")
    uploaded_vector = preprocess_uploaded_formula(excel_file)
    database_vectors, filenames = load_database_formulas()

    print("执行两步筛选匹配...")
    results = find_similar_formulas(
        uploaded_vector, database_vectors, filenames,
        JACCARD_THRESHOLD,
        WEIGHT_STRATEGY,
        TOP_N
    )

    if results:
        print(f"找到 {len(results)} 个相似配方")
        print_results(results)

        best_match = results[1]
        print(f"\n最佳匹配：{best_match['filename']} (相似度: {best_match['similarity_score']:.4f})")

        plot_comparison(uploaded_vector, best_match['db_vector'], best_match['filename'])
        visualize_jaccard_difference(uploaded_vector, best_match['db_vector'], best_match['filename'])
    else:
        print("未找到满足条件的相似配方，请降低杰卡德阈值")
