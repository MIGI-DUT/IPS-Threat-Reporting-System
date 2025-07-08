import pandas as pd
import json


def clean_envet_log(file_path):
    """
    清洗 envet_log JSON 数据，以便进行数据可视化。

    参数:
        file_path (str): envet_log JSON 文件的路径。

    返回:
        pandas.DataFrame: 一个已清洗的 DataFrame，可用于可视化。
    """
    # 明确指定编码为 'utf-8' 来打开文件，以解决 UnicodeDecodeError
    with open(file_path, 'r', encoding='utf-8') as f:
        # 加载整个 JSON 内容
        data = [json.loads(line) for line in f]

    df = pd.DataFrame(data)

    print("原始 DataFrame 信息:")
    df.info()
    print("\n原始 DataFrame 前几行:")
    print(df.head())

    # --- 1. 数据类型转换 ---

    # 将 'timestamp' 转换为 datetime 对象
    # 时间戳似乎是毫秒级的
    df['timestamp_ms'] = pd.to_numeric(df['timestamp'], errors='coerce')
    df['timestamp'] = pd.to_datetime(df['timestamp_ms'], unit='ms', errors='coerce')

    # 将数值字段转换为数值类型，强制转换错误会将无效解析转换为 NaN
    numerical_cols = ['reliability', 'severity', 'enrichments.dst_ip.malicious',
                      'enrichments.src_ip.malicious', 'number',
                      'enrichments.victim.in_range', 'original_reliability']
    for col in numerical_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # --- 2. 处理缺失值 ---

    # 为简单起见，我们用 'Unknown' 填充一些缺失的分类值
    # 或者，根据可视化需求，您可以删除包含关键数据缺失的行。
    categorical_cols_to_fill = ['src_ip_city', 'dst_ip_city', 'host', 'user_agent',
                                'status_msg', 'sub_category', 'classtype', 'kill_chain',
                                'intel_type', 'attack_status', 'tags', 'proto',
                                'dst_ip_country', 'victim_country_code']
    for col in categorical_cols_to_fill:
        if col in df.columns:
            df[col] = df[col].fillna('Unknown')

    # 对于数值列，填充 0 或平均值/中位数是可选项
    # 或者，如果它们是关键数据，删除行可能更好。
    # 这里，我们采用通用方法，将数值 NaN 填充为 0。
    for col in numerical_cols:
        if col in df.columns:
            df[col] = df[col].fillna(0)

    # 删除 'timestamp' 转换失败的行（如果有）
    df.dropna(subset=['timestamp'], inplace=True)

    # --- 3. 特征工程 ---

    # 从时间戳中提取小时、星期几和日期
    df['hour_of_day'] = df['timestamp'].dt.hour
    df['day_of_week'] = df['timestamp'].dt.day_name()
    df['event_date'] = df['timestamp'].dt.date

    # 如果存在 'dns' 字段且 'query' 是一个键，则从 'dns.query' 中提取域名
    if 'dns' in df.columns:
        df['dns_query'] = df['dns'].apply(lambda x: x.get('query') if isinstance(x, dict) else None)
        df['dns_qtype_name'] = df['dns'].apply(lambda x: x.get('qtype_name') if isinstance(x, dict) else None)
    else:
        df['dns_query'] = None
        df['dns_qtype_name'] = None

    # 如果存在 'desc' 字段，则从中提取信息
    if 'desc' in df.columns:
        # 示例: 从 desc 中解析 method, status_code, host, uri
        def parse_desc(desc):
            if not isinstance(desc, str):
                return None, None, None, None

            method = status_code = host = uri = None

            # 根据已知模式使用正则表达式或简单字符串分割
            if 'method:' in desc:
                method_match = desc.split('method:')[1].split('\n')[0].strip()
                method = method_match if method_match else None
            if 'status_code:' in desc:
                status_code_match = desc.split('status_code:')[1].split('\n')[0].strip()
                status_code = pd.to_numeric(status_code_match, errors='coerce') if status_code_match else None
            if 'host:' in desc:
                host_match = desc.split('host:')[1].split('\n')[0].strip()
                host = host_match if host_match else None
            if 'uri:' in desc:
                uri_match = desc.split('uri:')[1].split('\n')[0].strip()
                uri = uri_match if uri_match else None

            return method, status_code, host, uri

        df[['parsed_method', 'parsed_status_code', 'parsed_host', 'parsed_uri']] = df['desc'].apply(
            lambda x: pd.Series(parse_desc(x)))
        df['parsed_status_code'] = pd.to_numeric(df['parsed_status_code'], errors='coerce')  # 确保为数值类型

    # --- 4. 标准化分类数据 ---

    # 示例: 标准化国家代码
    # 这可能需要一个映射字典才能完全标准化
    country_mapping = {
        'China': 'CN',
        'CN': 'CN',
        # 根据需要添加更多映射
    }
    if 'dst_ip_country' in df.columns:
        df['dst_ip_country'] = df['dst_ip_country'].replace(country_mapping)
    if 'victim_country_code' in df.columns:
        df['victim_country_code'] = df['victim_country_code'].replace(country_mapping)

    # 将城市名称转换为一致的大小写（例如，首字母大写）
    for col in ['src_ip_city', 'dst_ip_city', 'victim_city']:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(lambda x: x.title() if x != 'Unknown' else x)

    # --- 5. 处理嵌套 JSON/扁平化（在 'dns' 和 'desc' 中已部分完成） ---
    # 如果需要特定的嵌套字段，可以进一步扁平化 'enrichments'
    # 例如，如果 'enrichments.src_ip.host_type' 是一个字典，您将从中提取。
    # 根据片段，它们似乎大部分已经扁平化。

    # 选择用于可视化的相关列，如果已解析则删除原始复杂列
    columns_to_keep = [
        'timestamp', 'timestamp_ms', 'event_date', 'hour_of_day', 'day_of_week',
        'src_ip', 'dst_ip', 'src_ip_city', 'dst_ip_city', 'dst_ip_country',
        'victim_city', 'victim_country_code', 'host', 'user_agent', 'status_msg',
        'reliability', 'severity', 'classtype', 'sub_category', 'kill_chain',
        'intel_type', 'attack_status', 'tags', 'proto', 'interface',
        'enrichments.dst_ip.malicious', 'enrichments.src_ip.malicious',
        'number', 'enrichments.victim.in_range', 'original_reliability',
        'dns_query', 'dns_qtype_name', 'parsed_method', 'parsed_status_code',
        'parsed_host', 'parsed_uri'
    ]

    # 过滤 DataFrame 中实际存在的列
    final_df_columns = [col for col in columns_to_keep if col in df.columns]
    final_df = df[final_df_columns].copy()

    print("\n清洗后的 DataFrame 信息:")
    final_df.info()
    print("\n清洗后的 DataFrame 前几行:")
    print(final_df.head())
    print("\n'classtype' 的值计数 (示例):")
    print(final_df['classtype'].value_counts())
    print("\n'parsed_method' 的值计数 (示例):")
    print(final_df['parsed_method'].value_counts())

    return final_df

# 运行清洗过程:
cleaned_data = clean_envet_log('../downloads/envet_log-20250707170407.json')
cleaned_data.to_csv('../downloads/cleaned_data.csv', index=False)

# 现在 'cleaned_data' DataFrame 已准备好用于 Matplotlib, Seaborn, Plotly 等可视化工具。
# 简单的绘图示例 (需要 matplotlib)
# import matplotlib.pyplot as plt
# if not cleaned_data.empty:
#     # 绘制事件随时间的变化
#     cleaned_data.set_index('timestamp').resample('H').size().plot(title='每小时事件数')
#     plt.ylabel('事件数量')
#     plt.show()

#     # classtype 分布的条形图
#     plt.figure(figsize=(10, 6))
#     cleaned_data['classtype'].value_counts().plot(kind='bar')
#     plt.title('Classtype 分布')
#     plt.xlabel('Classtype')
#     plt.ylabel('计数')
#     plt.xticks(rotation=45, ha='right')
#     plt.tight_layout()
#     plt.show()
