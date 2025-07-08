import pandas as pd

# 读取 CSV 文件
df = pd.read_csv('../downloads/cleaned_data.csv')

# 选择 dns_query 列不为空的行（既排除NaN也排除空字符串）
col = 'dns_query'
filtered_df = df[df[col].notna() & (df[col].astype(str).str.strip() != '')]

filtered_df = filtered_df.drop(columns=['src_ip_city', 'dst_ip_city', 'dst_ip_country', 'victim_city', 'victim_country_code',
                                        'sub_category', 'kill_chain', 'intel_type', 'tags', 'proto','interface', 'enrichments.dst_ip.malicious',
                                        'enrichments.src_ip.malicious', 'number', 'enrichments.victim.in_range', 'parsed_method',
                                        'parsed_status_code','parsed_host','parsed_uri'])
# 保存结果
filtered_df.to_csv('../downloads/filtered_data.csv', index=False)

# 查看前几行验证
print(filtered_df.head())
