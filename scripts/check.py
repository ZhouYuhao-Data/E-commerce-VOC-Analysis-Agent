import pandas as pd
import os

# ================= 配置区域 =================
# 确保文件名和你本地的一模一样
FILE_PATH = "日报（日用品）.xlsx"
OUTPUT_FILE = "日用品_中差评专项分析.xlsx"


def analyze_negative_reviews():
    print(f"正在读取文件: {FILE_PATH} ...")

    if not os.path.exists(FILE_PATH):
        print("❌ 错误：找不到文件，请确认文件名是否正确（后缀是 .xlsx）")
        return

    try:
        # 1. 精准读取：根据你的截图，表头在第 5 行 (header=4)
        df = pd.read_excel(FILE_PATH, header=9, engine='openpyxl')

        # 2. 验证列名：打印一下看对不对
        print("成功读取！检测到的列名：", list(df.columns[:5]), "...")

        # 3. 筛选核心列 (只取我们需要分析的)
        # 根据截图可以看到这些列名
        required_cols = ['日期', '商品名称', 'GMV', '销量', '中差评数', '中差评率']

        # 检查一下这些列是否存在，防止列名有微小空格差异
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            print(f"❌ 警告：没找到这些列 {missing_cols}，可能表头有空格，正在尝试自动修复...")
            # 自动去除列名两端的空格
            df.columns = [str(c).strip() for c in df.columns]

        # 再次尝试提取
        df_analysis = df[required_cols].copy()

        # 4. 数据清洗 (处理 NaN 和 格式)

        # 4.1 填充日期 (Forward Fill)：把“同属一天”的空行填上日期
        df_analysis['日期'] = df_analysis['日期'].ffill()

        # 4.2 去掉没有商品名称的行 (可能是汇总行或空行)
        df_analysis = df_analysis.dropna(subset=['商品名称'])
        # 去掉包含“汇总”字样的行
        df_analysis = df_analysis[~df_analysis['商品名称'].str.contains("汇总|合计", na=False)]

        # 4.3 数值处理
        # 中差评数：转为整数，空值填0
        df_analysis['中差评数'] = pd.to_numeric(df_analysis['中差评数'], errors='coerce').fillna(0).astype(int)

        # GMV：转为浮点数
        df_analysis['GMV'] = pd.to_numeric(df_analysis['GMV'], errors='coerce').fillna(0)

        # 中差评率：去掉百分号，转为小数 (如 "5.2%" -> 0.052)
        def clean_rate(x):
            if pd.isna(x): return 0.0
            x_str = str(x).replace('%', '').strip()
            try:
                return float(x_str) / 100 if '%' in str(x) else float(x_str)
            except:
                return 0.0

        # 注意：如果Excel里已经是小数格式（比如0.05），就不需要除以100，这里简单处理
        # 如果你的表里本来就是 0.05 显示为 5%，那就直接读数值即可。
        # 我们先按“可能是字符串带%”来处理，如果报错再调整。
        # 最稳妥的方式：直接信任Excel读取的数值，只做强制转换
        df_analysis['中差评率'] = pd.to_numeric(df_analysis['中差评率'].astype(str).str.replace('%', ''),
                                                errors='coerce')

        # ================= 5. 生成两个榜单 =================

        # 榜单 A：差评数量榜 (Top 20) - 这种商品正在疯狂得罪客户
        top_bad_count = df_analysis.sort_values(by='中差评数', ascending=False).head(20)

        # 榜单 B：差评率榜 (Top 20) - 这种商品可能销量不大，但卖一个骂一个 (排除销量太小的偶然数据)
        # 筛选条件：销量 > 10，避免只卖了1个且是差评导致100%差评率
        df_rate_valid = df_analysis[df_analysis['销量'] > 10]
        top_bad_rate = df_rate_valid.sort_values(by='中差评率', ascending=False).head(20)

        # ================= 6. 保存结果 =================
        print(f"正在保存分析结果到 {OUTPUT_FILE} ...")

        with pd.ExcelWriter(OUTPUT_FILE) as writer:
            df_analysis.to_excel(writer, sheet_name='清洗后总表', index=False)
            top_bad_count.to_excel(writer, sheet_name='差评数Top20(量大)', index=False)
            top_bad_rate.to_excel(writer, sheet_name='差评率Top20(质差)', index=False)

        print("🚀 大功告成！")
        print(f"请打开 {OUTPUT_FILE} 查看：")
        print("1. sheet '差评数Top20' -> 看看哪些爆款最近被骂得最惨")
        print("2. sheet '差评率Top20' -> 看看哪些品质量有硬伤")

    except Exception as e:
        print(f"❌ 发生未知错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    analyze_negative_reviews()