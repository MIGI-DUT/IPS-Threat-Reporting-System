import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体为黑体
plt.rcParams['axes.unicode_minus'] = False    # 正确显示负号

# 示例代码
plt.title('数量变化图')
plt.plot([1, 2, 3], [4, 5, 6])
plt.xlabel('时间')
plt.ylabel('数量')
plt.show()