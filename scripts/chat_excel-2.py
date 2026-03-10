import pandas as pd
from openai import OpenAI
import json
from tqdm import tqdm  # 显示处理进度条

# 1. 配置 API
client = OpenAI(
    api_key="",
    base_url=""
)


def analyze_feedback(text):
    if not text or pd.isna(text):
        return None

    # 将你的详细规则全部塞进 Prompt
    prompt = f"""
    你是一名专业的电商舆情分析专家，负责某头部电商的用户反馈分类。
    请根据【分类规则】分析用户评论，并严格按照指定【JSON格式】返回结果。

    【任务要求】
    1. 产品名称提取：识别名词性质的精简产品名称（去掉“自营”及修饰词），若无则填“无”。
    2. 一级/二级分类判定：根据评论主要意图，从规定分类中选择。
    3. 冲突处理：若评论涉及多个分类，请分析主要矛盾，归纳为最核心的一类。

    【分类规则】
    - 物流（二级分类：交付相关）：涉及快递、包装、派送、破损、冻坏、保鲜、运费、驿站等。
    - 客服（二级分类：客服相关）：涉及售后、退款、咨询、回复、态度、赔付等。
    - 直播间（二级分类：活动/优惠/专场/直播间）：涉及主播（如yoyo、顿顿）、直播内容、抽奖、福袋等。
    - 活动（二级分类：活动/优惠/专场/直播间）：涉及满减、积分、周边、甄果、会员价、券、优惠机制等。
    - 公关战略（二级分类：法务/公关/舆情相关）：涉及品牌形象、舆论、假货辟谣、公司团队、企业文化等。
    - 功能（二级分类：APP相关）：涉及APP界面、支付、搜索、按钮、版本更新、程序逻辑等。
    - 会员（二级分类：会员相关）：涉及会员权益、会员价等。
    - 产品（注意，产品需从以下三个二级分类中选一）：
        1. "产品开发"：之前没出过的产品（如：想要出方便面）。
        2. "产品建议"：已上线产品的优化建议（如：想要新口味、新包装）。
        3. "产品补货"：上线过但目前没货的（关键词：补货、回归、返场、下架、缺货）。

    【JSON输出格式】
    {{
        "product_name": "精简后的产品名称",
        "category_1": "一级分类名称",
        "category_2": "二级分类名称"
    }}

    用户评论内容："{text}"
    """

    # 后面的 API 调用逻辑保持不变...

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个只输出 JSON 格式的专业助手。"},
                {"role": "user", "content": prompt}
            ],
            response_format={'type': 'json_object'}  # 强制 AI 返回 JSON
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"处理失败: {text[:10]}... 错误: {e}")
        return None


# 3. 主程序逻辑
def main():
    input_file = "意见建议xlsx"  # 你的原始文件名
    output_file = "意见建议-分析结果.xlsx"

    # 读取 Excel
    print("正在读取数据...")
    df = pd.read_excel(input_file)

    # 假设你的评论在“摘要备注”这一列，请根据实际列名修改
    target_column = "具体内容"

    # 为了测试，我们可以先只跑前 10 行
    # df = df.head(10)

    results = []
    print("AI 正在分析中，请稍候...")

    # 使用 tqdm 显示进度
    for index, row in tqdm(df.iterrows(), total=len(df)):
        analysis = analyze_feedback(row[target_column])
        if analysis:
            results.append({
                "评价内容": row[target_column],
                "产品名称": analysis.get("product_name"),
                "一级分类": analysis.get("category_1"),
                "二级分类": analysis.get("category_2")
            })
        else:
            results.append({"评价内容": row[target_column], "产品名称": "分析失败", "一级分类": "", "二级分类": ""})

    # 4. 直接用 results 列表生成全新的 DataFrame
    # 此时 result_df 就只包含：原话、产品名称、一级分类、二级分类 这四列
    final_df = pd.DataFrame(results)

    # 5. 保存为新 Excel
    # 我们不再进行 concat 拼接，这样生成的就是一个全新的精简表
    final_df.to_excel(output_file, index=False)

    print(f"\n✅ 处理完成！已为您生成精简版汇总表: {output_file}")
    print(f"📊 表格字段：{', '.join(final_df.columns)}")


if __name__ == "__main__":
    main()