import pandas as pd

# ================= 配置区域 =================
# 填入你跑出来的两个结果文件的路径
FEEDBACK_FILE = '意见建议-分析结果.xlsx'  # 假设这是你第一张图的文件
DAILY_FILE = '日用品-分析结果.xlsx'  # 这是你第二张图(加了产品名称后)的文件

OUTPUT_TXT = '某头部电商_RAG知识库语料.txt'


def process_feedback_table():
    """处理用户建议反馈表"""
    try:
        df = pd.read_excel(FEEDBACK_FILE)
        texts = []
        for index, row in df.iterrows():
            product = str(row.get('产品名称', '未知商品'))
            content = str(row.get('评价内容', ''))
            cat1 = str(row.get('一级分类', ''))
            cat2 = str(row.get('二级分类', ''))

            # 拼接成自然语言
            text = f"【用户建议反馈】关于商品“{product}”的评价：用户反馈内容为“{content}”。该反馈属于{cat1}-{cat2}类目。"
            texts.append(text)
        return texts
    except Exception as e:
        print(f"读取 {FEEDBACK_FILE} 失败: {e}")
        return []


def process_daily_table():
    """处理日用品分析表"""
    try:
        df = pd.read_excel(DAILY_FILE)
        texts = []
        for index, row in df.iterrows():
            product = str(row.get('产品名称', '未知商品'))
            pain_point = str(row.get('核心痛点', ''))
            reason = str(row.get('责任归因', ''))
            sentiment = str(row.get('情感倾向', ''))
            action = str(row.get('改进建议', ''))

            # 过滤掉无效数据
            if product == '未知商品' and pain_point == '无':
                continue

            # 拼接成自然语言
            text = f"【日用品客诉分析】关于商品“{product}”的客诉记录：用户的核心痛点是“{pain_point}”，情感倾向为“{sentiment}”。经分析，该问题主要归因为“{reason}”。业务优化建议：“{action}”。"
            texts.append(text)
        return texts
    except Exception as e:
        print(f"读取 {DAILY_FILE} 失败: {e}")
        return []


def main():
    all_texts = []

    print("正在处理《用户建议反馈》...")
    all_texts.extend(process_feedback_table())

    print("正在处理《日用品-分析结果》...")
    all_texts.extend(process_daily_table())

    # 将所有文本写入 TXT 文件，用两个换行符隔开（方便 Dify 识别分块）
    print(f"正在生成知识库文件: {OUTPUT_TXT} ...")
    with open(OUTPUT_TXT, 'w', encoding='utf-8') as f:
        f.write('\n\n'.join(all_texts))

    print(f"🎉 成功！共生成 {len(all_texts)} 条 RAG 语料，可以直接上传到 Dify / Coze。")


if __name__ == "__main__":
    main()