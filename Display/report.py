import pandas as pd
from collections import Counter, defaultdict
import glob
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import io
import base64


class ThreatReportGenerator:
    def __init__(self):
        self.setup_fonts()
        self.styles = getSampleStyleSheet()
        self.create_custom_styles()

    def setup_fonts(self):
        """设置中文字体支持"""
        try:
            # 尝试注册中文字体
            # 您可能需要根据系统调整字体路径
            font_paths = [
                'C:/Windows/Fonts/simsun.ttc',  # Windows 宋体
                '/System/Library/Fonts/STHeiti Light.ttc',  # macOS 黑体
                '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',  # Linux
            ]

            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('SimSun', font_path))
                    break
        except:
            # 如果无法注册中文字体，使用默认字体
            print("警告：无法加载中文字体，将使用默认字体")

    def create_custom_styles(self):
        """创建自定义样式"""
        try:
            font_name = 'SimSun'
        except:
            font_name = 'Helvetica'

        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Title'],
            fontSize=18,
            spaceAfter=30,
            fontName=font_name,
            alignment=1  # 居中对齐
        )

        self.heading_style = ParagraphStyle(
            'CustomHeading',
            parent=self.styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
            fontName=font_name,
            textColor=colors.darkblue
        )

        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=10,
            fontName=font_name,
            spaceAfter=6
        )

    def find_log_file(self):
        """查找日志文件"""
        files = glob.glob("../downloads/*envet_log*.xlsx")
        if not files:
            raise FileNotFoundError("未找到匹配 'envet_log*.xlsx' 的文件")
        return files[0]

    def load_data(self, file_path):
        """加载数据"""
        try:
            df = pd.read_excel(file_path, header=0, engine='openpyxl')
            print(f"成功加载文件: {file_path}")
            print(f"数据行数: {len(df)}")
            return df
        except Exception as e:
            raise Exception(f"加载数据失败: {str(e)}")

    def preprocess_data(self, df):
        """数据预处理"""
        # 转换时间列
        time_column = '发现时间'
        if time_column in df.columns:
            try:
                df[time_column] = pd.to_datetime(df[time_column], format='%Y-%m-%d %H:%M:%S')
            except:
                # 尝试其他时间格式
                df[time_column] = pd.to_datetime(df[time_column], infer_datetime_format=True)

        # 判断是否是客户端源IP
        def is_client_ip(ip):
            ip_str = str(ip)
            return ip_str.startswith('172.') or ip_str.startswith('192.')

        # 添加IP类型列
        if '源IP' in df.columns:
            df['IP类型'] = df['源IP'].apply(lambda ip: '客户端' if is_client_ip(ip) else '服务端')

        return df

    def analyze_threats(self, df):
        """威胁分析"""
        # 定义列名映射
        column_mapping = {
            'threat_category': '威胁类别',
            'threat_name': '威胁名称',
            'severity': '威胁等级',
            'src_ip': '源IP',
            'dst_ip': '目的IP',
            'dst_port': '目的端口',
            'protocol': '应用层协议',
            'time': '发现时间'
        }

        # 验证列存在性
        existing_columns = {eng: col for eng, col in column_mapping.items() if col in df.columns}

        threat_stats = {
            'total_events': len(df),
            'threat_categories': {},
            'threat_names': {},
            'severity_levels': {},
            'source_ips': {},
            'destination_ips': {},
            'time_distribution': defaultdict(int),
            'top_malicious_ips': {},
            'common_ports': {},
            'protocols': {},
            'client_analysis': {},
            'server_analysis': {}
        }

        # 基本统计
        for eng, col in existing_columns.items():
            if col in df.columns:
                if eng == 'threat_category':
                    threat_stats['threat_categories'] = df[col].value_counts().to_dict()
                elif eng == 'threat_name':
                    threat_stats['threat_names'] = df[col].value_counts().to_dict()
                elif eng == 'severity':
                    threat_stats['severity_levels'] = df[col].value_counts().to_dict()
                elif eng == 'src_ip':
                    threat_stats['source_ips'] = df[col].value_counts().to_dict()
                elif eng == 'dst_ip':
                    threat_stats['destination_ips'] = df[col].value_counts().to_dict()
                elif eng == 'dst_port':
                    threat_stats['common_ports'] = df[col].value_counts().head(10).to_dict()
                elif eng == 'protocol':
                    threat_stats['protocols'] = df[col].value_counts().to_dict()

        # 时间分布分析
        if '发现时间' in df.columns:
            for time in df['发现时间']:
                if pd.notna(time):
                    try:
                        hour = time.hour
                        threat_stats['time_distribution'][hour] += 1
                    except:
                        pass

        # 客户端和服务端分析
        if 'IP类型' in df.columns and '源IP' in df.columns:
            client_df = df[df['IP类型'] == '客户端']
            server_df = df[df['IP类型'] == '服务端']

            threat_stats['client_analysis'] = self.analyze_top_ips(client_df)
            threat_stats['server_analysis'] = self.analyze_top_ips(server_df)

        return threat_stats

    def analyze_top_ips(self, sub_df):
        """分析TOP IP"""
        if sub_df.empty or '源IP' not in sub_df.columns:
            return {}

        top_ips = sub_df['源IP'].value_counts().head(5)
        ip_analysis = {}

        for ip, count in top_ips.items():
            ip_rows = sub_df[sub_df['源IP'] == ip]

            # 统计威胁等级和名称
            threat_stat = {}
            if '威胁等级' in ip_rows.columns and '威胁名称' in ip_rows.columns:
                combined = zip(ip_rows['威胁等级'], ip_rows['威胁名称'])
                threat_counter = Counter([f"[{level}] {name}" for level, name in combined])
                threat_stat = dict(threat_counter)

            ip_analysis[ip] = {
                'count': count,
                'threats': threat_stat
            }

        return ip_analysis

    def create_charts(self, threat_stats):
        """创建图表"""
        # 设置中文字体
        plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False

        chart_files = []

        # 1. 威胁类别分布图
        if threat_stats['threat_categories']:
            plt.figure(figsize=(10, 6))
            categories = list(threat_stats['threat_categories'].keys())
            values = list(threat_stats['threat_categories'].values())

            bars = plt.bar(categories, values, color='skyblue', edgecolor='navy', alpha=0.7)
            plt.title('威胁类别分布', fontsize=14, fontweight='bold')
            plt.ylabel('事件数量', fontsize=12)
            plt.xlabel('威胁类别', fontsize=12)
            plt.xticks(rotation=45, ha='right')

            # 添加数值标签
            for bar, value in zip(bars, values):
                plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.1,
                         str(value), ha='center', va='bottom')

            plt.tight_layout()
            chart_file = 'threat_categories.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight')
            plt.close()
            chart_files.append(chart_file)

        # 2. 时间分布图
        if threat_stats['time_distribution']:
            plt.figure(figsize=(12, 6))
            hours = sorted(threat_stats['time_distribution'].items())
            x, y = zip(*hours) if hours else ([], [])

            plt.plot(x, y, marker='o', linewidth=2, markersize=6, color='red')
            plt.fill_between(x, y, alpha=0.3, color='red')
            plt.title('威胁事件时间分布', fontsize=14, fontweight='bold')
            plt.xlabel('小时', fontsize=12)
            plt.ylabel('事件数量', fontsize=12)
            plt.xticks(range(24))
            plt.grid(True, alpha=0.3)
            plt.tight_layout()

            chart_file = 'time_distribution.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight')
            plt.close()
            chart_files.append(chart_file)

        # 3. 严重程度分布饼图
        if threat_stats['severity_levels']:
            plt.figure(figsize=(8, 8))
            labels = list(threat_stats['severity_levels'].keys())
            sizes = list(threat_stats['severity_levels'].values())
            colors_list = ['red', 'orange', 'yellow', 'green', 'blue'][:len(labels)]

            plt.pie(sizes, labels=labels, colors=colors_list, autopct='%1.1f%%',
                    startangle=90, textprops={'fontsize': 10})
            plt.title('威胁严重程度分布', fontsize=14, fontweight='bold')
            plt.axis('equal')

            chart_file = 'severity_distribution.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight')
            plt.close()
            chart_files.append(chart_file)

        return chart_files

    def create_pdf_report(self, threat_stats, chart_files, output_file='threat_report.pdf'):
        """创建PDF报告"""
        doc = SimpleDocTemplate(output_file, pagesize=A4)
        story = []

        # 标题
        story.append(Paragraph("网络安全威胁分析报告", self.title_style))
        story.append(Spacer(1, 20))

        # 报告信息
        report_info = f"""
        <b>生成时间:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>
        <b>总事件数:</b> {threat_stats['total_events']}<br/>
        <b>分析维度:</b> 威胁类别、严重程度、源IP、时间分布等
        """
        story.append(Paragraph(report_info, self.normal_style))
        story.append(Spacer(1, 20))

        # 1. 威胁概览
        story.append(Paragraph("1. 威胁概览", self.heading_style))

        severity_text = "威胁等级分布:<br/>"
        for level, count in threat_stats['severity_levels'].items():
            percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
            severity_text += f"• {level}: {count} 起 ({percentage:.1f}%)<br/>"

        story.append(Paragraph(severity_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 2. 威胁类别分析
        story.append(Paragraph("2. 威胁类别分析", self.heading_style))

        category_text = ""
        for category, count in list(threat_stats['threat_categories'].items())[:10]:
            percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
            category_text += f"• {category}: {count} 起 ({percentage:.1f}%)<br/>"

        story.append(Paragraph(category_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 3. 常见威胁类型
        story.append(Paragraph("3. 常见威胁类型", self.heading_style))

        threat_text = ""
        for name, count in list(threat_stats['threat_names'].items())[:10]:
            threat_text += f"• {name}: {count} 起<br/>"

        story.append(Paragraph(threat_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 4. 主要威胁源分析
        story.append(Paragraph("4. 主要威胁源分析", self.heading_style))

        source_text = f"涉及源IP总数: {len(threat_stats['source_ips'])}<br/><br/>"
        source_text += "TOP 5 威胁源IP:<br/>"

        top_sources = sorted(threat_stats['source_ips'].items(), key=lambda x: x[1], reverse=True)[:5]
        for ip, count in top_sources:
            source_text += f"• {ip}: {count} 起<br/>"

        story.append(Paragraph(source_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 5. 客户端威胁分析
        if threat_stats['client_analysis']:
            story.append(Paragraph("5. 客户端威胁分析", self.heading_style))

            client_text = "前五频发的客户端源IP及其威胁统计:<br/><br/>"
            for ip, data in threat_stats['client_analysis'].items():
                client_text += f"<b>IP: {ip}</b> (出现 {data['count']} 次)<br/>"
                for threat, count in data['threats'].items():
                    client_text += f"  • {threat}: {count}<br/>"
                client_text += "<br/>"

            story.append(Paragraph(client_text, self.normal_style))
            story.append(Spacer(1, 15))

        # 6. 服务端威胁分析
        if threat_stats['server_analysis']:
            story.append(Paragraph("6. 服务端威胁分析", self.heading_style))

            server_text = "前五频发的服务端源IP及其威胁统计:<br/><br/>"
            for ip, data in threat_stats['server_analysis'].items():
                server_text += f"<b>IP: {ip}</b> (出现 {data['count']} 次)<br/>"
                for threat, count in data['threats'].items():
                    server_text += f"  • {threat}: {count}<br/>"
                server_text += "<br/>"

            story.append(Paragraph(server_text, self.normal_style))
            story.append(Spacer(1, 15))

        # 7. 时间分布分析
        if threat_stats['time_distribution']:
            story.append(Paragraph("7. 威胁时间分布", self.heading_style))

            time_text = "24小时威胁事件分布:<br/>"
            sorted_hours = sorted(threat_stats['time_distribution'].items())
            for hour, count in sorted_hours:
                time_text += f"• {hour:02d}:00-{hour + 1:02d}:00: {count} 起<br/>"

            story.append(Paragraph(time_text, self.normal_style))
            story.append(Spacer(1, 15))

        # 8. 协议和端口分析
        story.append(Paragraph("8. 协议与端口分析", self.heading_style))

        proto_text = "常用协议分布:<br/>"
        for proto, count in list(threat_stats['protocols'].items())[:10]:
            proto_text += f"• {proto}: {count} 次<br/>"

        proto_text += "<br/>常见目标端口:<br/>"
        for port, count in list(threat_stats['common_ports'].items())[:10]:
            proto_text += f"• 端口 {port}: {count} 次<br/>"

        story.append(Paragraph(proto_text, self.normal_style))
        story.append(PageBreak())

        # 添加图表
        story.append(Paragraph("9. 数据可视化", self.heading_style))

        for chart_file in chart_files:
            if os.path.exists(chart_file):
                try:
                    story.append(Image(chart_file, width=6 * inch, height=4 * inch))
                    story.append(Spacer(1, 20))
                except Exception as e:
                    print(f"无法添加图表 {chart_file}: {e}")

        # 生成PDF
        doc.build(story)

        # 清理临时图表文件
        for chart_file in chart_files:
            if os.path.exists(chart_file):
                try:
                    os.remove(chart_file)
                except:
                    pass

        return output_file

    def generate_report(self, output_file='threat_report.pdf'):
        """生成完整报告"""
        try:
            # 1. 查找日志文件
            log_file = self.find_log_file()
            print(f"📄 找到日志文件: {log_file}")

            # 2. 加载数据
            df = self.load_data(log_file)

            # 3. 数据预处理
            df = self.preprocess_data(df)

            # 4. 威胁分析
            threat_stats = self.analyze_threats(df)

            # 5. 创建图表
            chart_files = self.create_charts(threat_stats)

            # 6. 生成PDF报告
            pdf_file = self.create_pdf_report(threat_stats, chart_files, output_file)

            print(f"✅ PDF报告已生成: {pdf_file}")
            return pdf_file

        except Exception as e:
            print(f"❌ 报告生成失败: {str(e)}")
            raise


if __name__ == "__main__":
    # 创建报告生成器实例
    generator = ThreatReportGenerator()

    # 生成报告
    try:
        report_file = generator.generate_report('网络安全威胁分析报告.pdf')
        print(f"\n🎉 报告生成成功！")
        print(f"📄 文件位置: {report_file}")
        print(f"📊 报告包含: 威胁统计、IP分析、时间分布、可视化图表等")
    except Exception as e:
        print(f"\n❌ 程序执行失败: {str(e)}")
        print("请检查:")
        print("1. Excel文件是否存在于 ../downloads/ 目录")
        print("2. 文件名是否包含 'envet_log'")
        print("3. 数据格式是否正确")
        print("4. 是否已安装必要的库: pip install pandas openpyxl matplotlib reportlab")