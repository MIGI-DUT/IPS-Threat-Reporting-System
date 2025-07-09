import pandas as pd
from collections import defaultdict
from datetime import datetime
import matplotlib.pyplot as plt


def generate_threat_report(log_file):
    # 读取日志文件
    try:
        # 尝试不同的编码方式读取Excel文件
        df = pd.read_excel(log_file, header=0)  # 确保使用第一行作为列名
    except Exception as e:
        return f"Error reading log file: {str(e)}"

    # 检查列名是否正确，打印出来看看
    print("检测到的列名:", df.columns.tolist())

    # 尝试匹配中文列名或英文列名
    time_column = None
    if '发现时间' in df.columns:
        time_column = '发现时间'
    elif 'Detection Time' in df.columns:  # 如果有英文列名
        time_column = 'Detection Time'

    if time_column is None:
        return "错误：在Excel文件中找不到时间列（可能是'发现时间'或'Detection Time'）"

    # 数据清洗和预处理
    df = df.dropna(how='all')

    # 转换时间列
    try:
        df[time_column] = pd.to_datetime(df[time_column])
    except Exception as e:
        return f"时间列转换错误: {str(e)}"

    # 分析威胁数据
    threat_stats = analyze_threats(df, time_column)

    # 生成报告
    report = create_report(threat_stats, df, time_column)

    return report


def analyze_threats(df, time_column):
    # 定义列名映射（中英文对照）
    column_mapping = {
        'threat_category': ['威胁类别', 'Threat Category'],
        'threat_name': ['威胁名称', 'Threat Name'],
        'severity': ['威胁等级', 'Severity Level'],
        'src_ip': ['源IP', 'Source IP'],
        'dst_ip': ['目的IP', 'Destination IP'],
        'dst_port': ['目的端口', 'Destination Port'],
        'protocol': ['应用层协议', 'Application Protocol']
    }

    # 查找实际列名
    actual_columns = {}
    for key, possible_names in column_mapping.items():
        for name in possible_names:
            if name in df.columns:
                actual_columns[key] = name
                break

    # 威胁类别统计
    threat_stats = {
        'total_events': len(df),
        'threat_categories': df[actual_columns['threat_category']].value_counts().to_dict(),
        'threat_names': df[actual_columns['threat_name']].value_counts().to_dict(),
        'severity_levels': df[actual_columns['severity']].value_counts().to_dict(),
        'source_ips': df[actual_columns['src_ip']].value_counts().to_dict(),
        'destination_ips': df[actual_columns['dst_ip']].value_counts().to_dict(),
        'time_distribution': defaultdict(int),
        'top_malicious_ips': {},
        'common_ports': df[actual_columns['dst_port']].value_counts().head(10).to_dict(),
        'protocols': df[actual_columns['protocol']].value_counts().to_dict()
    }

    # 按小时统计事件分布
    for time in df[time_column]:
        hour = time.hour
        threat_stats['time_distribution'][hour] += 1

    # 统计恶意IP
    if 'threat-intelligence-alarm' in threat_stats['threat_categories']:
        malicious_ips = df[df[actual_columns['threat_category']] == 'threat-intelligence-alarm'][
            actual_columns['dst_ip']].value_counts().head(5)
        threat_stats['top_malicious_ips'] = malicious_ips.to_dict()

    return threat_stats


def create_report(threat_stats, df, time_column):
    # 报告头部
    report = f"""
    ==================== 网络安全威胁报告 ====================
    生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    分析时间段: {df[time_column].min()} 至 {df[time_column].max()}
    总事件数: {threat_stats['total_events']}
    ========================================================

    """

    # 威胁概览
    report += "1. 威胁概览\n"
    report += f"- 检测到的高危威胁(severity_4): {threat_stats['severity_levels'].get('severity_4', 0)} 起\n"
    report += f"- 检测到的中危威胁(severity_3): {threat_stats['severity_levels'].get('severity_3', 0)} 起\n"
    report += f"- 检测到的低危威胁(severity_1-2): {threat_stats['severity_levels'].get('severity_1', 0) + threat_stats['severity_levels'].get('severity_2', 0)} 起\n\n"

    # 威胁类别分析
    report += "2. 威胁类别分析\n"
    for category, count in threat_stats['threat_categories'].items():
        report += f"- {category}: {count} 起 ({count / threat_stats['total_events']:.1%})\n"
    report += "\n"

    # 常见威胁类型
    report += "3. 常见威胁类型\n"
    for name, count in list(threat_stats['threat_names'].items())[:5]:
        report += f"- {name}: {count} 起\n"
    report += "\n"

    # 源IP分析
    report += "4. 主要威胁源分析\n"
    report += f"- 涉及源IP总数: {len(threat_stats['source_ips'])}\n"
    top_sources = sorted(threat_stats['source_ips'].items(), key=lambda x: x[1], reverse=True)[:5]
    for ip, count in top_sources:
        report += f"- {ip}: {count} 起\n"
    report += "\n"

    # 恶意目的IP
    if threat_stats['top_malicious_ips']:
        report += "5. 主要恶意目的IP\n"
        for ip, count in threat_stats['top_malicious_ips'].items():
            report += f"- {ip}: {count} 次查询\n"
        report += "\n"

    # 时间分布
    report += "6. 威胁时间分布\n"
    sorted_hours = sorted(threat_stats['time_distribution'].items())
    for hour, count in sorted_hours:
        report += f"- {hour:02d}:00-{hour + 1:02d}:00: {count} 起\n"
    report += "\n"

    # 协议和端口分析
    report += "7. 协议与端口分析\n"
    report += "- 常用协议:\n"
    for proto, count in threat_stats['protocols'].items():
        report += f"  - {proto}: {count} 次\n"

    report += "\n- 常见目标端口:\n"
    for port, count in threat_stats['common_ports'].items():
        report += f"  - 端口 {port}: {count} 次\n"

    # 关键发现和建议
    report += """

    8. 关键发现与建议
    - 主要威胁类型是恶意域名DNS查询和Tor网络流量
    - 建议对频繁出现的内网IP进行深入调查
    - 检测到多个远程控制工具的使用，建议检查是否授权
    - 检测到VPN和代理工具使用，建议审查网络策略
    - 建议加强DNS查询监控
    """

    return report


def save_visualizations(threat_stats, output_dir="."):
    """生成可视化图表"""
    # 威胁类别分布
    plt.figure(figsize=(10, 6))
    pd.Series(threat_stats['threat_categories']).plot(kind='bar')
    plt.title('威胁类别分布')
    plt.ylabel('事件数量')
    plt.tight_layout()
    plt.savefig(f"{output_dir}/threat_categories.png")
    plt.close()

    # 时间分布
    plt.figure(figsize=(10, 6))
    hours = sorted(threat_stats['time_distribution'].items())
    x, y = zip(*hours)
    plt.plot(x, y, marker='o')
    plt.title('威胁事件时间分布')
    plt.xlabel('小时')
    plt.ylabel('事件数量')
    plt.xticks(range(24))
    plt.grid()
    plt.tight_layout()
    plt.savefig(f"{output_dir}/time_distribution.png")
    plt.close()


# 使用示例
if __name__ == "__main__":
    log_file = "envet_log-20250707173129.xlsx"
    report = generate_threat_report(log_file)

    # 保存报告
    with open("threat_report.txt", "w", encoding="utf-8") as f:
        f.write(report)

    # 生成可视化图表
    df = pd.read_excel(log_file, header=0)
    time_column = '发现时间' if '发现时间' in df.columns else 'Detection Time'
    threat_stats = analyze_threats(df, time_column)
    save_visualizations(threat_stats)

    print("威胁报告已生成: threat_report.txt")
    print("可视化图表已保存: threat_categories.png 和 time_distribution.png")