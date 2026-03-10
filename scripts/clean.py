import pandas as pd
import os
import warnings

# 忽略 openpyxl 的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ================= 1. 配置区域 =================
FILES = {
    "weekly": "日报（周维度）.xlsx",
    "daily": "日报（日用品）.xlsx",
    "fresh": "日报（果蔬生鲜）.xlsx"
}
OUTPUT_FILE = "经营分析总表_清洗版.xlsx"


# ================= 2. 核心逻辑函数 =================

def safe_contains(val, target_list):
    """安全检查：确保只在字符串中搜索关键词"""
    if pd.isna(val): return False
    val_str = str(val)
    return any(t in val_str for t in target_list)


def clean_excel_data(file_path, label):
    if not os.path.exists(file_path):
        print(f"⚠️ 跳过：找不到文件 {file_path}")
        return pd.DataFrame()

    print(f"正在读取并分析: {file_path} ...")

    try:
        # 1. 读取原始数据 (不带表头)
        raw_df = pd.read_excel(file_path, header=None, engine='openpyxl')

        # 2. 智能定位表头行 (加固版)
        header_idx = -1
        for i, row in raw_df.head(100).iterrows():
            # 将该行所有单元格转为字符串并检查
            row_content = [str(x) for x in row.values if pd.notna(x)]
            has_gmv = any("GMV" in s or "成交" in s for s in row_content)
            has_refund = any("赔付" in s or "退款" in s or "金额" in s for s in row_content)

            if has_gmv and has_refund:
                header_idx = i
                print(f"✅ [{label}] 在第 {i} 行锁定表头")
                break

        if header_idx == -1:
            print(f"❌ [{label}] 未找到有效表头。")
            return pd.DataFrame()

        # 3. 提取并清理
        df = raw_df.iloc[header_idx:].copy()
        df.columns = df.iloc[0]  # 设置第一行为列名
        df = df.iloc[1:].copy()  # 移除表头重复行

        # 4. 清理列名
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]

        # 5. 识别核心列
        gmv_col = next((c for c in df.columns if "GMV" in c or "成交" in c), None)
        refund_col = next((c for c in df.columns if "赔付" in c), None)

        if not gmv_col: return pd.DataFrame()

        # 6. 处理堆叠表格 (剔除中间重复的表头行)
        # 只要这一行的 GMV 列内容依然包含 "GMV" 字样，就说明它是重复表头
        df = df[~df[gmv_col].astype(str).str.contains("GMV|成交|金额", na=False)]

        # 剔除汇总行
        first_col = df.columns[0]
        df = df[~df[first_col].astype(str).str.contains("汇总|Total|合计", na=False)]

        # 7. 填充合并单元格 (日期/类目)
        df.iloc[:, 0:2] = df.iloc[:, 0:2].ffill()

        # 8. 强制数值化 (处理 ￥, 逗号, 以及可能的错误字符)
        def to_num(x):
            if pd.isna(x): return 0
            s = str(x).replace('¥', '').replace(',', '').strip()
            try:
                return float(s)
            except:
                return 0

        df[gmv_col] = df[gmv_col].apply(to_num)
        if refund_col:
            df[refund_col] = df[refund_col].apply(to_num)
            df = df.rename(columns={refund_col: '赔付金额'})

        df = df.rename(columns={gmv_col: 'GMV'})

        # 9. 计算赔付率
        if '赔付金额' in df.columns:
            df['赔付率'] = df.apply(lambda x: x['赔付金额'] / x['GMV'] if x['GMV'] > 0 else 0, axis=1)

        df['数据来源'] = label
        return df

    except Exception as e:
        print(f"❌ 处理 {label} 时发生错误: {e}")
        import traceback
        traceback.print_exc()  # 打印具体哪一行报错，方便调试
        return pd.DataFrame()


# ================= 3. 主程序 =================

if __name__ == "__main__":
    results = []
    for k, v in [('weekly', '周维度'), ('daily', '日用品'), ('fresh', '果蔬生鲜')]:
        clean_df = clean_excel_data(FILES[k], v)
        if not clean_df.empty:
            results.append(clean_df)

    if results:
        final_df = pd.concat(results, axis=0, ignore_index=True, sort=False)
        # 统一输出列
        cols = ['数据来源', '日期', '商品名称', '三级类目', 'GMV', '赔付金额', '赔付率']
        existing_cols = [c for c in cols if c in final_df.columns]
        other_cols = [c for c in final_df.columns if c not in existing_cols]
        final_df[existing_cols + other_cols].to_excel(OUTPUT_FILE, index=False)
        print(f"\n🚀 【完美输出】文件已保存至：{OUTPUT_FILE}")
    else:
        print("❌ 最终未生成任何数据，请检查 Excel 内是否存在 GMV 和 赔付 关键字。")