import pandas as pd
import json
import time
from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# ================= 配置区域 =================
# 建议使用 DeepSeek 或 Gemini 的 API，这里以通用 OpenAI SDK 格式为例
API_KEY = ""
BASE_URL = ""  # 或者 Gemini 的 Base URL
MODEL_NAME = "deepseek-chat"  # 或者 gemini-1.5-pro

client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

# 读取 Excel 文件 (根据你的截图文件名)
INPUT_FILE = '日用品.xlsx'
OUTPUT_FILE = '日用品-分析结果.xlsx'


# ================= 核心处理逻辑 =================

# 1. 定义带有重试机制的 API 调用函数 (体现工程稳定性)
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def analyze_row(row_data):
    """
    调用 LLM 进行单行数据分析
    """
    category = row_data['诉求分类3级名称']
    content = row_data['摘要备注']

    # 如果备注太短或无效，直接跳过，节省 Token
    if pd.isna(content) or len(str(content)) < 2:
        return None

    prompt = f"""
            你是一个电商分析师。请分析以下客服记录：
            预设分类：{category}
            用户备注：{content}

            请提取并返回如下 JSON 格式（不要Markdown）：
            {{
                "产品名称": "从备注中提取具体商品名(如'纯棉卫衣')，如果没有明确提及具体商品，填'未知商品'",
                "核心痛点": "简短描述具体问题",
                "责任归因": "从[供应链, 物流, 运营, 系统, 用户]中选一个",
                "情感倾向": "负面/中性/正面",
                "改进建议": "给业务方的具体建议"
            }}
            """

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,  # 低温度保证输出稳定
            response_format={"type": "json_object"}  # 强制 JSON (如果模型支持)
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"Error analyzing row: {e}")
        raise e  # 抛出异常触发 @retry 重试


# ================= 主程序 =================

def main():
    # 1. 读取数据
    print(f"正在读取 {INPUT_FILE}...")
    df = pd.read_excel(INPUT_FILE)

    # 截取你需要处理的列，防止数据量太大
    # 假设我们只处理前 50 条测试，正式跑可以去掉 .head()
    target_data = (df[['诉求分类3级名称', '摘要备注']].copy() .head(200))

    results = []

    # 2. 并发处理 (体现高性能)
    # 这是一个典型的 I/O 密集型任务，使用 ThreadPoolExecutor 提速
    max_workers = 5  # 根据 API 的 Rate Limit 调整并发数

    print(f"开始 AI 分析，并发线程数: {max_workers}...")

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交任务
        future_to_index = {
            executor.submit(analyze_row, row): index
            for index, row in target_data.iterrows()
        }

        # 使用 tqdm 显示进度条
        for future in tqdm(as_completed(future_to_index), total=len(target_data)):
            index = future_to_index[future]
            try:
                data = future.result()
                if data:
                    data['原始行号'] = index
                    results.append(data)
            except Exception as e:
                print(f"行 {index} 处理最终失败: {e}")

        # ... (前面的代码保持不变)

    # 3. 结果合并与保存
    print("正在合并数据...")
    results_df = pd.DataFrame(results)

    # 关键修改：只保留原始表中的两列 + AI 生成的所有列
    # 确保 '原始行号' 存在以便对齐

    # 从原始 df 中只提取我们需要的两列
    original_columns_kept = df.loc[results_df['原始行号'], ['诉求分类3级名称', '摘要备注']].reset_index(drop=True)

    # 将 AI 结果（去掉 '原始行号' 列，因为它只是个索引）与 原始两列 横向拼接
    final_output = pd.concat([original_columns_kept, results_df.drop(columns=['原始行号'])], axis=1)

    print(f"正在保存至 {OUTPUT_FILE}...")
    final_output.to_excel(OUTPUT_FILE, index=False)
    print("任务完成！🚀")

if __name__ == "__main__":
    main()