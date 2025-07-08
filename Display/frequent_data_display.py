import pandas as pd
from collections import Counter
import glob

# === Step 1: 找到 ../downloads/ 目录中包含 'event_log' 的 Excel 文件 ===
files = glob.glob("../downloads/*envet_log*.xlsx")
if not files:
    print("❌ ../downloads/ 目录下未找到匹配 'envet_log*.xlsx' 的文件。")
    exit()

file_path = files[0]
print(f"📄 正在读取文件: {file_path}")

# === Step 2: 读取数据 ===
df = pd.read_excel(file_path)

# 判断是否是客户端源IP
def is_client_ip(ip):
    ip_str = str(ip)
    return ip_str.startswith('172.') or ip_str.startswith('192.')

# 添加 IP 类型列
df['IP类型'] = df['源IP'].apply(lambda ip: '客户端' if is_client_ip(ip) else '服务端')

# === Step 3: 分别处理客户端和服务端 ===

def analyze_top_ips(sub_df, label):
    print(f"\n🔍 前五频发的{label}源IP及其[威胁等级+名称]统计：")
    top_ips = sub_df['源IP'].value_counts().head(5)
    for ip, count in top_ips.items():
        ip_rows = sub_df[sub_df['源IP'] == ip]
        combined = zip(ip_rows['威胁等级'], ip_rows['威胁名称'])
        threat_stat = Counter([f"[{level}] {name}" for level, name in combined])

        print(f"\n📌 IP: {ip} （出现 {count} 次）")
        for key, c in threat_stat.items():
            print(f" - {key}: {c}")

# 客户端分析
client_df = df[df['IP类型'] == '客户端']
analyze_top_ips(client_df, "客户端")

# 服务端分析
server_df = df[df['IP类型'] == '服务端']
analyze_top_ips(server_df, "服务端")
