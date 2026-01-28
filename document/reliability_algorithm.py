import pandas as pd
import numpy as np

# 1. 核心参数配置（基于星能江夏数值）
CONFIG = {
    "故障率": {"电缆": 0.09282879, "架空": 0.15337829},
    "预安排停电率": {"电缆": 0.0221, "架空": 0.0656},
    "故障修复时间_r": 3.073,  # 小时
    "隔离时间_自动化": 0.557,  # 小时
    "隔离时间_非自动化": 2.0,  # 小时
    "预安排停电时间_rs": 5.475  # 小时
}


def customized_reliability_algorithm(excel_path):
    # 读取主线和分支 Sheet
    with pd.ExcelFile(excel_path) as xls:
        df_main = pd.read_excel(xls, '主线')
        df_branch = pd.read_excel(xls, '分支')

    # 计算全局总用户数 N (算法分母)
    # 对应表头：用户数量(台)
    total_users = df_main['用户数量(台)'].sum() + df_branch['用户数量(台)'].sum()

    # --- 处理主线逻辑 ---
    def process_main(df):
        # 识别敷设方式：从“线路型号”中模糊匹配
        df['is_cable'] = df['线路型号'].apply(lambda x: "YJV" in str(x) or "YJLV" in str(x))
        df['lambda_f'] = df['is_cable'].apply(lambda x: CONFIG["故障率"]["电缆"] if x else CONFIG["故障率"]["架空"])
        df['lambda_s'] = df['is_cable'].apply(
            lambda x: CONFIG["预安排停电率"]["电缆"] if x else CONFIG["预安排停电率"]["架空"])

        # 识别自动化状态：对应“起点是否自动化”
        df['r_iso'] = df['起点是否自动化'].apply(
            lambda x: CONFIG["隔离时间_自动化"] if "TRUE" in str(x).upper() else CONFIG["隔离时间_非自动化"]
        )

        # 计算 SAIDI/SAIFI 贡献
        # L * lambda * (r+iso) * n / N
        df['saidi_f_cont'] = (df['lambda_f'] * df['长度(km)'] * (CONFIG["故障修复时间_r"] + df['r_iso']) * df[
            '用户数量(台)']) / total_users
        df['saidi_s_cont'] = (df['lambda_s'] * df['长度(km)'] * CONFIG["预安排停电时间_rs"] * df[
            '用户数量(台)']) / total_users
        df['saifi_f_cont'] = (df['lambda_f'] * df['长度(km)'] * df['用户数量(台)']) / total_users
        df['saifi_s_cont'] = (df['lambda_s'] * df['长度(km)'] * df['用户数量(台)']) / total_users
        return df

    # --- 处理分支逻辑 ---
    def process_branch(df):
        # 识别敷设方式
        df['is_cable'] = df['线路型号'].apply(lambda x: "YJV" in str(x) or "YJLV" in str(x))
        df['lambda_f'] = df['is_cable'].apply(lambda x: CONFIG["故障率"]["电缆"] if x else CONFIG["故障率"]["架空"])
        df['lambda_s'] = df['is_cable'].apply(
            lambda x: CONFIG["预安排停电率"]["电缆"] if x else CONFIG["预安排停电率"]["架空"])

        # 识别自动化状态：对应“是否自动化”
        df['r_iso'] = df['是否自动化'].apply(
            lambda x: CONFIG["隔离时间_自动化"] if "TRUE" in str(x).upper() else CONFIG["隔离时间_非自动化"]
        )

        # 计算贡献
        df['saidi_f_cont'] = (df['lambda_f'] * df['长度(km)'] * (CONFIG["故障修复时间_r"] + df['r_iso']) * df[
            '用户数量(台)']) / total_users
        df['saidi_s_cont'] = (df['lambda_s'] * df['长度(km)'] * CONFIG["预安排停电时间_rs"] * df[
            '用户数量(台)']) / total_users
        df['saifi_f_cont'] = (df['lambda_f'] * df['长度(km)'] * df['用户数量(台)']) / total_users
        df['saifi_s_cont'] = (df['lambda_s'] * df['长度(km)'] * df['用户数量(台)']) / total_users
        return df

    # 执行计算
    main_processed = process_main(df_main)
    branch_processed = process_branch(df_branch)

    # 汇总计算结果
    def summarize(df_list, label):
        combined = pd.concat(df_list)
        f_saidi = combined['saidi_f_cont'].sum()
        s_saidi = combined['saidi_s_cont'].sum()
        total_saidi = f_saidi + s_saidi
        total_saifi = combined['saifi_f_cont'].sum() + combined['saifi_s_cont'].sum()

        return {
            "分类": label,
            "总用户数": int(combined['用户数量(台)'].sum()),
            "总长度(km)": round(combined['长度(km)'].sum(), 3),
            "SAIDI (h/户·年)": round(total_saidi, 4),
            "SAIFI (次/户·年)": round(total_saifi, 4),
            "ASAI (%)": round((1 - total_saidi / 8760) * 100, 5)
        }

    results = [
        summarize([main_processed], "主线路汇总"),
        summarize([branch_processed], "分支线路汇总"),
        summarize([main_processed, branch_processed], "全线路汇总")
    ]

    return pd.DataFrame(results)

# 使用方式
final_report = customized_reliability_algorithm(r"D:\works\电网\人工智能\AI需求\配网全景拓扑\需求\技术调研\供电可靠性\算法\10kV安54新窑线.xlsx")
print(final_report)