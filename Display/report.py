import pandas as pd
from collections import Counter, defaultdict
import glob
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from matplotlib.font_manager import FontProperties
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
import matplotlib.font_manager as fm
import os
import io
import base64
import numpy as np

plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体为黑体
plt.rcParams['axes.unicode_minus'] = False  # 正确显示负号


class EnhancedThreatReportGenerator:  # 确保这一行存在
    def __init__(self):
        self.setup_fonts()
        self.setup_colors()
        self.styles = getSampleStyleSheet()
        self.create_custom_styles()
        self.setup_matplotlib_style()

    def setup_fonts(self):
        """设置中文字体支持"""
        try:
            # 尝试注册中文字体
            font_paths = [
                'C:/Windows/Fonts/simsun.ttc',  # Windows 宋体
                'C:/Windows/Fonts/simhei.ttf',  # Windows 黑体
            ]

            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('SimSun', font_path))
                    break
        except:
            print("警告：无法加载中文字体，将使用默认字体")

    def setup_colors(self):
        """设置颜色主题"""
        self.colors = {
            'primary': colors.Color(0.1, 0.2, 0.5),  # 深蓝色
            'secondary': colors.Color(0.8, 0.1, 0.1),  # 深红色
            'accent': colors.Color(0.2, 0.7, 0.3),  # 绿色
            'warning': colors.Color(0.9, 0.6, 0.1),  # 橙色
            'light_gray': colors.Color(0.95, 0.95, 0.95),  # 浅灰色
            'dark_gray': colors.Color(0.3, 0.3, 0.3),  # 深灰色
            'white': colors.white,
            'black': colors.black
        }

    def setup_matplotlib_style(self):
        font_path = 'C:/Windows/Fonts/simsun.ttc'  # 或 simhei.ttf
        if os.path.exists(font_path):
            self.font_prop = FontProperties(fname=font_path)
            plt.rcParams['font.sans-serif'] = [self.font_prop.get_name()]
            print(f"✅ matplotlib字体设置为: {self.font_prop.get_name()}")
        else:
            self.font_prop = None
            print("⚠️ 未找到中文字体，可能会乱码")

        plt.rcParams['axes.unicode_minus'] = False
        sns.set_style("whitegrid")
        sns.set_palette("husl")

    def create_custom_styles(self):
        """创建自定义样式"""
        try:
            font_name = 'SimSun'
        except:
            font_name = 'Helvetica'

        # 主标题样式
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Title'],
            fontSize=24,
            spaceAfter=30,
            spaceBefore=20,
            fontName=font_name,
            alignment=1,  # 居中对齐
            textColor=self.colors['primary'],
            borderWidth=2,
            borderColor=self.colors['primary'],
            borderPadding=10
        )

        # 章节标题样式
        self.heading_style = ParagraphStyle(
            'CustomHeading',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=12,
            spaceBefore=20,
            fontName=font_name,
            textColor=self.colors['primary'],
            borderWidth=1,
            borderColor=self.colors['light_gray'],
            borderPadding=5,
            backColor=self.colors['light_gray']
        )

        # 子标题样式
        self.subheading_style = ParagraphStyle(
            'CustomSubheading',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceAfter=8,
            spaceBefore=12,
            fontName=font_name,
            textColor=self.colors['secondary'],
            leftIndent=10
        )

        # 正文样式
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=11,
            fontName=font_name,
            spaceAfter=8,
            leftIndent=10,
            rightIndent=10,
            leading=14
        )

        # 重要信息样式
        self.highlight_style = ParagraphStyle(
            'HighlightStyle',
            parent=self.styles['Normal'],
            fontSize=11,
            fontName=font_name,
            spaceAfter=8,
            leftIndent=10,
            rightIndent=10,
            backColor=colors.lightyellow,
            borderColor=self.colors['warning'],
            borderWidth=1,
            borderPadding=8
        )

        # 警告样式
        self.warning_style = ParagraphStyle(
            'WarningStyle',
            parent=self.styles['Normal'],
            fontSize=11,
            fontName=font_name,
            spaceAfter=8,
            leftIndent=10,
            rightIndent=10,
            backColor=colors.lightcoral,
            borderColor=self.colors['secondary'],
            borderWidth=1,
            borderPadding=8,
            textColor=colors.darkred
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
            'daily_distribution': defaultdict(int),
            'top_malicious_ips': {},
            'common_ports': {},
            'protocols': {},
            'client_analysis': {},
            'server_analysis': {},
            'risk_score': 0
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
                        day = time.strftime('%Y-%m-%d')
                        threat_stats['time_distribution'][hour] += 1
                        threat_stats['daily_distribution'][day] += 1
                    except:
                        pass

        # 客户端和服务端分析
        if 'IP类型' in df.columns and '源IP' in df.columns:
            client_df = df[df['IP类型'] == '客户端']
            server_df = df[df['IP类型'] == '服务端']

            threat_stats['client_analysis'] = self.analyze_top_ips(client_df)
            threat_stats['server_analysis'] = self.analyze_top_ips(server_df)

        # 计算风险评分
        threat_stats['risk_score'] = self.calculate_risk_score(threat_stats)

        return threat_stats

    def calculate_risk_score(self, stats):
        """计算风险评分 (0-100)"""
        score = 0

        # 基于威胁等级的评分
        severity_weights = {'高': 30, '中': 20, '低': 10}
        for level, count in stats['severity_levels'].items():
            weight = severity_weights.get(level, 15)
            score += min(count * weight / 100, 30)

        # 基于威胁事件总数的评分
        total_events = stats['total_events']
        if total_events > 1000:
            score += 25
        elif total_events > 500:
            score += 15
        elif total_events > 100:
            score += 10

        # 基于威胁源IP数量的评分
        unique_ips = len(stats['source_ips'])
        if unique_ips > 50:
            score += 20
        elif unique_ips > 20:
            score += 15
        elif unique_ips > 10:
            score += 10

        return min(score, 100)

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

    def create_enhanced_charts(self, threat_stats):
        """创建增强的图表"""
        chart_files = []

        fp = self.font_prop

        # 1. 威胁类别分布图 - 美化版
        if threat_stats['threat_categories']:
            fig, ax = plt.subplots(figsize=(12, 8))

            categories = list(threat_stats['threat_categories'].keys())
            values = list(threat_stats['threat_categories'].values())

            # 使用渐变色
            colors = plt.cm.viridis(np.linspace(0, 1, len(categories)))

            bars = ax.bar(categories, values, color=colors, edgecolor='white', linewidth=2)

            # 添加阴影效果
            for bar in bars:
                bar.set_alpha(0.8)

            ax.set_title('威胁类别分布', fontsize=18, fontweight='bold', pad=20, fontproperties=fp)
            ax.set_ylabel('事件数量', fontsize=14, fontproperties=fp)
            ax.set_xlabel('威胁类别', fontsize=14, fontproperties=fp)
            ax.tick_params(axis='x', rotation=45)

            # 添加数值标签
            for bar, value in zip(bars, values):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                        f'{value:,}', ha='center', va='bottom', fontweight='bold')

            # 添加网格
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            ax.set_axisbelow(True)

            plt.tight_layout()
            chart_file = 'threat_categories_enhanced.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            chart_files.append(chart_file)

        # 2. 时间分布热力图 - 修复下面图表显示问题
        if threat_stats['time_distribution'] and threat_stats['daily_distribution']:
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))

            # 小时分布
            hours = sorted(threat_stats['time_distribution'].items())
            x, y = zip(*hours) if hours else ([], [])

            ax1.plot(x, y, marker='o', linewidth=3, markersize=8, color='#FF6B6B', alpha=0.8)
            ax1.fill_between(x, y, alpha=0.3, color='#FF6B6B')
            ax1.set_title('24小时威胁事件分布', fontsize=16, fontweight='bold', fontproperties=fp)
            ax1.set_xlabel('小时', fontsize=12, fontproperties=fp)
            ax1.set_ylabel('事件数量', fontsize=12, fontproperties=fp)
            ax1.set_xticks(range(24))
            ax1.grid(True, alpha=0.3)

            # 设置x轴刻度标签字体
            for label in ax1.get_xticklabels():
                label.set_fontproperties(fp)
            for label in ax1.get_yticklabels():
                label.set_fontproperties(fp)

            # 日期分布 - 修改条件判断，即使只有一天数据也显示
            if len(threat_stats['daily_distribution']) >= 1:
                daily_data = sorted(threat_stats['daily_distribution'].items())
                dates, counts = zip(*daily_data)

                ax2.bar(range(len(dates)), counts, color='#4ECDC4', alpha=0.7)
                ax2.set_title('日期威胁事件分布', fontsize=16, fontweight='bold', fontproperties=fp)
                ax2.set_xlabel('日期', fontsize=12, fontproperties=fp)
                ax2.set_ylabel('事件数量', fontsize=12, fontproperties=fp)

                # 设置x轴标签
                if len(dates) > 10:
                    step = max(1, len(dates) // 10)
                    ax2.set_xticks(range(0, len(dates), step))
                    ax2.set_xticklabels([dates[i] for i in range(0, len(dates), step)], rotation=45)
                else:
                    # 如果日期数量少，显示所有日期
                    ax2.set_xticks(range(len(dates)))
                    ax2.set_xticklabels(dates, rotation=45)

                ax2.grid(True, alpha=0.3)

                # 设置坐标轴刻度标签字体
                for label in ax2.get_xticklabels():
                    label.set_fontproperties(fp)
                for label in ax2.get_yticklabels():
                    label.set_fontproperties(fp)
            else:
                # 如果没有日期数据，显示提示信息
                ax2.text(0.5, 0.5, '暂无日期分布数据', ha='center', va='center',
                         transform=ax2.transAxes, fontsize=14, fontproperties=fp)
                ax2.set_title('日期威胁事件分布', fontsize=16, fontweight='bold', fontproperties=fp)
                ax2.set_xlabel('日期', fontsize=12, fontproperties=fp)
                ax2.set_ylabel('事件数量', fontsize=12, fontproperties=fp)

            plt.tight_layout()
            chart_file = 'time_distribution_enhanced.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            chart_files.append(chart_file)

        # 3. 威胁严重程度饼图 - 修复版本
        if threat_stats['severity_levels']:
            fig, ax = plt.subplots(figsize=(12, 10))  # 增大图表尺寸

            labels = list(threat_stats['severity_levels'].keys())
            sizes = list(threat_stats['severity_levels'].values())

            # 定义颜色映射 - 使用更鲜明的颜色
            color_map = {
                '高': '#FF4444',  # 鲜红色
                '中': '#FFA500',  # 橙色
                '低': '#32CD32',  # 绿色
                '严重': '#8B0000',  # 深红色
                '警告': '#FF8C00',  # 深橙色
                '信息': '#4169E1',  # 蓝色
                '提示': '#9932CC'  # 紫色
            }

            # 为每个标签分配颜色，如果没有预定义颜色则使用默认色盘
            colors_list = []
            default_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', '#FF9FF3', '#54A0FF']

            for i, label in enumerate(labels):
                if label in color_map:
                    colors_list.append(color_map[label])
                else:
                    colors_list.append(default_colors[i % len(default_colors)])

            # 设置饼图参数，不显示标签和百分比
            wedges, texts = ax.pie(sizes,
                                   labels=None,  # 不显示标签
                                   colors=colors_list,
                                   autopct=None,  # 不显示百分比
                                   startangle=90,
                                   explode=[0.1 if label == '高' else 0.05 for label in labels])

            ax.set_title('威胁严重程度分布', fontsize=20, fontweight='bold', pad=30, fontproperties=fp)

            # 创建图例标签，避免使用可能显示为方格的字符
            legend_labels = []
            for label, size in zip(labels, sizes):
                percentage = (size / sum(sizes)) * 100
                # 使用英文字符替代可能有问题的中文字符
                legend_labels.append(f'{label}等级: {size}个 ({percentage:.1f}%)')

            # 调整图例位置，避免重叠
            legend = ax.legend(wedges, legend_labels,
                               title="威胁等级统计",
                               loc="center left",
                               bbox_to_anchor=(1.2, 0.5),
                               fontsize=12,
                               title_fontsize=14)

            # 设置图例标题字体
            if fp:
                legend.get_title().set_fontproperties(fp)
                # 设置图例文本字体
                for text in legend.get_texts():
                    text.set_fontproperties(fp)

            # 确保图表布局合理
            plt.subplots_adjust(left=0.1, right=0.75)

            chart_file = 'severity_distribution_enhanced.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            chart_files.append(chart_file)

        # 4. TOP IP威胁分析图
        if threat_stats['source_ips']:
            fig, ax = plt.subplots(figsize=(12, 8))

            top_ips = sorted(threat_stats['source_ips'].items(), key=lambda x: x[1], reverse=True)[:10]
            ips, counts = zip(*top_ips)

            bars = ax.barh(range(len(ips)), counts, color='#FF7F7F', alpha=0.8)

            ax.set_yticks(range(len(ips)))
            ax.set_yticklabels(ips)
            ax.set_xlabel('威胁事件数量', fontsize=12, fontproperties=fp)
            ax.set_title('TOP 10 威胁源IP', fontsize=16, fontweight='bold', fontproperties=fp)

            # 添加数值标签
            for i, (bar, count) in enumerate(zip(bars, counts)):
                ax.text(bar.get_width() + max(counts) * 0.01, bar.get_y() + bar.get_height() / 2,
                        f'{count:,}', ha='left', va='center', fontweight='bold')

            ax.grid(axis='x', alpha=0.3)
            plt.tight_layout()

            chart_file = 'top_ips_enhanced.png'
            plt.savefig(chart_file, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            chart_files.append(chart_file)

        return chart_files

    def create_summary_table(self, threat_stats):
        """创建汇总表格"""
        data = [
            ['指标', '数值', '描述'],
            ['总威胁事件', f"{threat_stats['total_events']:,}", '检测到的威胁事件总数'],
            ['威胁类别数', f"{len(threat_stats['threat_categories'])}", '涉及的威胁类别种类'],
            ['威胁源IP数', f"{len(threat_stats['source_ips'])}", '产生威胁的源IP数量'],
            ['风险评分', f"{threat_stats['risk_score']:.1f}/100", '综合风险评估分数'],
        ]

        # 添加威胁等级统计
        for level, count in threat_stats['severity_levels'].items():
            data.append([f'{level}等级威胁', f"{count:,}", f'{level}等级威胁事件数量'])

        table = Table(data, colWidths=[2 * inch, 1.5 * inch, 3 * inch])
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'SimSun'),  # 表头用宋体
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),

            ('FONTNAME', (0, 1), (-1, -1), 'SimSun'),  # 内容区也设置为宋体
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.colors['light_gray']]),
        ]))

        return table

    def create_risk_indicator(self, risk_score):
        """创建风险指示器"""
        if risk_score >= 70:
            risk_level = "高风险"
            risk_color = self.colors['secondary']
            risk_desc = "需要立即采取安全措施"
        elif risk_score >= 40:
            risk_level = "中风险"
            risk_color = self.colors['warning']
            risk_desc = "建议加强安全监控"
        else:
            risk_level = "低风险"
            risk_color = self.colors['accent']
            risk_desc = "当前安全状况良好"

        risk_text = f"""
        <b><font color="{risk_color}">风险等级: {risk_level}</font></b><br/>
        <b>风险评分: {risk_score:.1f}/100</b><br/>
        {risk_desc}
        """

        return risk_text, risk_color

    def create_pdf_report(self, threat_stats, chart_files, output_file='enhanced_threat_report.pdf'):
        """创建美化的PDF报告"""
        doc = SimpleDocTemplate(output_file, pagesize=A4, topMargin=1 * inch, bottomMargin=1 * inch)
        story = []

        # 封面
        story.append(Paragraph("网络安全威胁分析报告", self.title_style))
        story.append(Spacer(1, 30))

        # 报告基本信息
        report_info = f"""
        <b>生成时间:</b> {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}<br/>
        <b>分析时间段:</b> 全量数据分析<br/>
        <b>报告状态:</b> <font color="green">已完成</font>
        """
        story.append(Paragraph(report_info, self.highlight_style))
        story.append(Spacer(1, 20))

        # 风险评估摘要
        risk_text, risk_color = self.create_risk_indicator(threat_stats['risk_score'])
        story.append(Paragraph("🔍 风险评估摘要", self.heading_style))
        story.append(
            Paragraph(risk_text, self.warning_style if threat_stats['risk_score'] >= 70 else self.highlight_style))
        story.append(Spacer(1, 20))

        # 威胁概览表格
        story.append(Paragraph("📊 威胁统计概览", self.heading_style))
        summary_table = self.create_summary_table(threat_stats)
        story.append(summary_table)
        story.append(Spacer(1, 20))

        # 1. 威胁类别分析
        story.append(Paragraph("1. 威胁类别分析", self.heading_style))

        if threat_stats['threat_categories']:
            category_text = "本次分析共发现以下威胁类别:<br/><br/>"
            for i, (category, count) in enumerate(list(threat_stats['threat_categories'].items())[:10], 1):
                percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
                icon = "🔴" if percentage > 20 else "🟡" if percentage > 10 else "🟢"
                category_text += f"{icon} <b>{category}</b>: {count:,} 起 ({percentage:.1f}%)<br/>"

            story.append(Paragraph(category_text, self.normal_style))

        story.append(Spacer(1, 15))

        # 2. 威胁等级分布
        story.append(Paragraph("2. 威胁等级分布", self.heading_style))

        severity_text = "威胁等级统计分析:<br/><br/>"
        for level, count in threat_stats['severity_levels'].items():
            percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
            icon = "🚨" if level == "高" else "⚠️" if level == "中" else "ℹ️"
            color = "red" if level == "高" else "orange" if level == "中" else "green"
            severity_text += f'{icon} <font color="{color}"><b>{level}等级威胁</b></font>: {count:,} 起 ({percentage:.1f}%)<br/>'

        story.append(Paragraph(severity_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 3. 常见威胁类型TOP 10
        story.append(Paragraph("3. 常见威胁类型 TOP 10", self.heading_style))

        if threat_stats['threat_names']:
            threat_text = ""
            for i, (name, count) in enumerate(list(threat_stats['threat_names'].items())[:10], 1):
                medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else f"{i}."
                threat_text += f"{medal} <b>{name}</b>: {count:,} 起<br/>"

            story.append(Paragraph(threat_text, self.normal_style))

        story.append(Spacer(1, 15))

        # 4. 威胁源分析
        story.append(Paragraph("4. 威胁源分析", self.heading_style))

        source_text = f"🌐 <b>威胁源IP统计</b><br/>"
        source_text += f"• 涉及源IP总数: <b>{len(threat_stats['source_ips']):,}</b><br/>"
        source_text += f"• 平均每IP威胁数: <b>{threat_stats['total_events'] / len(threat_stats['source_ips']):.1f}</b><br/><br/>"

        source_text += "🔝 <b>TOP 5 威胁源IP:</b><br/>"
        top_sources = sorted(threat_stats['source_ips'].items(), key=lambda x: x[1], reverse=True)[:5]
        for i, (ip, count) in enumerate(top_sources, 1):
            source_text += f"{i}. <b>{ip}</b>: {count:,} 起<br/>"

        story.append(Paragraph(source_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 5. 客户端威胁分析
        if threat_stats['client_analysis']:
            story.append(Paragraph("5. 客户端威胁分析", self.heading_style))

            client_text = "🖥️ <b>客户端威胁统计</b><br/><br/>"
            client_text += "以下是检测到威胁活动最频繁的客户端IP:<br/><br/>"

            for i, (ip, data) in enumerate(threat_stats['client_analysis'].items(), 1):
                client_text += f"<b>{i}. IP: {ip}</b> (共 {data['count']} 次威胁)<br/>"
                for threat, count in list(data['threats'].items())[:3]:  # 只显示前3个
                    client_text += f"  • {threat}: {count} 次<br/>"
                client_text += "<br/>"

            story.append(Paragraph(client_text, self.normal_style))
            story.append(Spacer(1, 15))

            # 6. 服务端威胁分析
        if threat_stats['server_analysis']:
            story.append(Paragraph("6. 服务端威胁分析", self.heading_style))

            server_text = "🖥️ <b>服务端威胁统计</b><br/><br/>"
            server_text += "以下是检测到威胁活动最频繁的服务端IP:<br/><br/>"

            for i, (ip, data) in enumerate(threat_stats['server_analysis'].items(), 1):
                server_text += f"<b>{i}. IP: {ip}</b> (共 {data['count']} 次威胁)<br/>"
                for threat, count in list(data['threats'].items())[:3]:  # 只显示前3个
                    server_text += f"  • {threat}: {count} 次<br/>"
                server_text += "<br/>"

            story.append(Paragraph(server_text, self.normal_style))
            story.append(Spacer(1, 15))

            # 7. 时间分布分析
        if threat_stats['time_distribution']:
            story.append(Paragraph("7. 威胁时间分布", self.heading_style))

            time_text = "⏰ <b>24小时威胁事件分布</b><br/><br/>"
            sorted_hours = sorted(threat_stats['time_distribution'].items())

            # 计算峰值时间
            peak_hour = max(sorted_hours, key=lambda x: x[1]) if sorted_hours else (0, 0)
            time_text += f"🔝 <b>威胁峰值时间:</b> {peak_hour[0]:02d}:00-{peak_hour[0] + 1:02d}:00 ({peak_hour[1]} 起)<br/><br/>"

            # 分时段统计
            time_ranges = {
                '深夜(00:00-06:00)': sum(count for hour, count in sorted_hours if 0 <= hour < 6),
                '早晨(06:00-12:00)': sum(count for hour, count in sorted_hours if 6 <= hour < 12),
                '下午(12:00-18:00)': sum(count for hour, count in sorted_hours if 12 <= hour < 18),
                '晚间(18:00-24:00)': sum(count for hour, count in sorted_hours if 18 <= hour < 24),
            }

            for period, count in time_ranges.items():
                percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
                time_text += f"• {period}: {count:,} 起 ({percentage:.1f}%)<br/>"

            story.append(Paragraph(time_text, self.normal_style))
            story.append(Spacer(1, 15))

            # 8. 协议与端口分析
        story.append(Paragraph("8. 协议与端口分析", self.heading_style))

        proto_text = "🌐 <b>网络协议分布</b><br/><br/>"
        if threat_stats['protocols']:
            for i, (proto, count) in enumerate(list(threat_stats['protocols'].items())[:10], 1):
                percentage = (count / threat_stats['total_events']) * 100 if threat_stats['total_events'] > 0 else 0
                proto_text += f"{i}. <b>{proto}</b>: {count:,} 次 ({percentage:.1f}%)<br/>"

        proto_text += "<br/>🔌 <b>常见目标端口</b><br/><br/>"
        if threat_stats['common_ports']:
            for i, (port, count) in enumerate(list(threat_stats['common_ports'].items())[:10], 1):
                # 常见端口描述
                port_desc = {
                    '80': 'HTTP',
                    '443': 'HTTPS',
                    '22': 'SSH',
                    '21': 'FTP',
                    '25': 'SMTP',
                    '53': 'DNS',
                    '3389': 'RDP',
                    '445': 'SMB',
                    '993': 'IMAPS',
                    '995': 'POP3S'
                }
                desc = port_desc.get(str(port), '未知服务')
                proto_text += f"{i}. <b>端口 {port}</b> ({desc}): {count:,} 次<br/>"

        story.append(Paragraph(proto_text, self.normal_style))
        story.append(Spacer(1, 15))

        # 9. 安全建议
        story.append(Paragraph("9. 安全建议", self.heading_style))

        recommendations = []

        # 基于风险评分的建议
        if threat_stats['risk_score'] >= 70:
            recommendations.append("🚨 <b>紧急建议</b>：系统风险评分较高，建议立即进行全面安全检查")
            recommendations.append("🔒 启动应急响应流程，隔离高风险IP地址")
        elif threat_stats['risk_score'] >= 40:
            recommendations.append("⚠️ <b>中级建议</b>：加强安全监控，定期检查威胁状态")
        else:
            recommendations.append("✅ <b>基础建议</b>：继续维持当前安全措施")

        # 基于威胁类型的建议
        if '高' in threat_stats['severity_levels'] and threat_stats['severity_levels']['高'] > 0:
            recommendations.append("🔥 针对高级威胁，建议更新防护策略和规则")

        # 基于时间分布的建议
        if threat_stats['time_distribution']:
            peak_times = sorted(threat_stats['time_distribution'].items(), key=lambda x: x[1], reverse=True)[:3]
            peak_hours = [str(h[0]) for h in peak_times]
            recommendations.append(f"🕐 加强 {', '.join(peak_hours)} 时段的安全监控")

        # 基于IP数量的建议
        if len(threat_stats['source_ips']) > 50:
            recommendations.append("🌐 威胁源IP数量较多，建议实施IP地址黑名单策略")

        recommendations.append("📋 定期更新威胁情报和安全规则")
        recommendations.append("🎯 对高频威胁IP进行深度分析和追踪")
        recommendations.append("📊 建立长期威胁监控和趋势分析机制")

        rec_text = "<br/>".join(recommendations)
        story.append(Paragraph(rec_text, self.highlight_style))
        story.append(PageBreak())

        # 10. 数据可视化
        story.append(Paragraph("10. 数据可视化", self.heading_style))
        story.append(Paragraph("以下图表展示了威胁数据的详细分析结果:", self.normal_style))
        story.append(Spacer(1, 20))

        # 添加图表
        for i, chart_file in enumerate(chart_files, 1):
            if os.path.exists(chart_file):
                try:
                    # 根据图表类型添加标题
                    chart_titles = {
                        'threat_categories_enhanced.png': f'图表 {i}: 威胁类别分布统计',
                        'time_distribution_enhanced.png': f'图表 {i}: 威胁时间分布分析',
                        'severity_distribution_enhanced.png': f'图表 {i}: 威胁严重程度分布',
                        'top_ips_enhanced.png': f'图表 {i}: TOP 10 威胁源IP分析'
                    }

                    chart_title = chart_titles.get(os.path.basename(chart_file), f'图表 {i}')
                    story.append(Paragraph(chart_title, self.subheading_style))
                    story.append(Spacer(1, 10))

                    # 添加图表
                    story.append(Image(chart_file, width=6.5 * inch, height=4.5 * inch))
                    story.append(Spacer(1, 20))
                except Exception as e:
                    print(f"无法添加图表 {chart_file}: {e}")

        # 11. 报告总结
        story.append(Paragraph("11. 报告总结", self.heading_style))

        summary_text = f"""
                    <b>📈 数据概览:</b><br/>
                    • 本次分析共处理威胁事件 <b>{threat_stats['total_events']:,}</b> 起<br/>
                    • 涉及威胁类别 <b>{len(threat_stats['threat_categories'])}</b> 种<br/>
                    • 威胁源IP地址 <b>{len(threat_stats['source_ips'])}</b> 个<br/>
                    • 系统风险评分 <b>{threat_stats['risk_score']:.1f}/100</b><br/><br/>

                    <b>🎯 关键发现:</b><br/>
                    • 最活跃的威胁类别: <b>{list(threat_stats['threat_categories'].keys())[0] if threat_stats['threat_categories'] else 'N/A'}</b><br/>
                    • 最频繁的威胁源IP: <b>{list(threat_stats['source_ips'].keys())[0] if threat_stats['source_ips'] else 'N/A'}</b><br/>
                    • 威胁活动峰值时间: <b>{max(threat_stats['time_distribution'].items(), key=lambda x: x[1])[0] if threat_stats['time_distribution'] else 'N/A'}:00</b><br/><br/>

                    <b>📋 后续行动:</b><br/>
                    • 持续监控高风险IP和威胁类别<br/>
                    • 定期更新安全策略和防护规则<br/>
                    • 加强团队安全意识培训<br/>
                    • 建立完善的威胁响应机制
                    """

        story.append(Paragraph(summary_text, self.normal_style))
        story.append(Spacer(1, 20))

        # 报告结束标识
        end_text = f"""
                    <b>--- 报告结束 ---</b><br/>
                    <i>报告生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}</i><br/>
                    """
        story.append(Paragraph(end_text, self.normal_style))

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

    def generate_report(self, output_file='enhanced_threat_report.pdf'):
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

            # 5. 创建增强图表
            chart_files = self.create_enhanced_charts(threat_stats)

            # 6. 生成PDF报告
            pdf_file = self.create_pdf_report(threat_stats, chart_files, output_file)

            print(f"✅ 增强版PDF报告已生成: {pdf_file}")
            return pdf_file

        except Exception as e:
            print(f"❌ 报告生成失败: {str(e)}")
            raise


if __name__ == "__main__":
    generator = EnhancedThreatReportGenerator()
    try:
        report_file = generator.generate_report('网络安全威胁分析报告.pdf')
        print(f"\n🎉 增强版报告生成成功！")
        print(f"📄 文件位置: {report_file}")
        print(f"📊 报告包含: 威胁统计、IP分析、时间分布、美化图表、安全建议等")
        print(f"✨ 新增功能: 风险评分、可视化增强、详细建议、报告总结")
    except Exception as e:
        print(f"\n❌ 程序执行失败: {str(e)}")
        print("请检查:")
        print("1. Excel文件是否存在于 ../downloads/ 目录")
        print("2. 文件名是否包含 'envet_log'")
        print("3. 数据格式是否正确")
        print("4. 是否已安装必要的库: pip install pandas openpyxl matplotlib seaborn reportlab")
