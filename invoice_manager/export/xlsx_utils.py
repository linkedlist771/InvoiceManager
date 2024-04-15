import pandas as pd
import shutil
import os


def filter_invoice_by_year_month(df: pd.DataFrame, year: int, month: int):
    def check_date(_date: str, _year: int, _month: int) -> bool:
        _date = str(_date)
        _year = str(_year)
        _month = str(_month)
        year_, month_, _ = _date.split("-")
        # strip the zero on the left side of the month
        month_ = month_.lstrip("0")
        return year_ == _year and month_ == _month

    column_name = "凭证日期（开票）"
    invoice_date = df[column_name]
    # check if the date is in the year and month
    matched_idx = [check_date(date, year, month) for date in invoice_date]
    return df[matched_idx]


def get_deviceid_by_invoice_id(invoice_id: str, df: pd.DataFrame, _device_info_df: pd.DataFrame) -> str:
    # invoice_id unused, incase we need to use it
    item_ids = df["测试项目编号匹配"]
    # random choose one for it
    item_id = item_ids.iloc[0]
    matched_idx = _device_info_df["测试编号"].str.contains(item_id)
    clean_matched_idx = matched_idx.fillna(False)
    device_id = _device_info_df[clean_matched_idx]["仪器单位资产管理编号"].values[0]
    return device_id


def get_company_name_by_invoice_id(invoice_id: str, df: pd.DataFrame, _custom_contact_locations_df: pd.DataFrame) -> str:
    # invoice_id unused, incase we need to use it
    item_ids = df["客户全称"]
    # random choose one for it
    item_id = item_ids.iloc[0]
    return item_id


def get_custom_info_by_invoice_id(invoice_id: str, df: pd.DataFrame, _custom_contact_locations_df: pd.DataFrame) -> dict:
    # invoice_id unused, incase we need to use it
    item_id = get_company_name_by_invoice_id(invoice_id, df, _custom_contact_locations_df)
    # item_id is substr of "测试编号“ column in _device_info_df
    # _device_info_df = _device_info_df.dropna(subset=["设备编号"])
    matched_idx = _custom_contact_locations_df["委托单位"] == item_id
    clean_matched_idx = matched_idx.fillna(False)
    custom_info = _custom_contact_locations_df[clean_matched_idx].iloc[0].to_dict()
    return custom_info


def get_operator_and_billing(invoice_id: str, device_id: str, _operator_billing_df: pd.DataFrame) -> dict:
    # operator_billing_df 仪器单位资产管理编号
    matched_idx = _operator_billing_df["仪器单位资产管理编号"] == device_id
    clean_matched_idx = matched_idx.fillna(False)
    # 返回第一个的字典格式吧
    operator_info = _operator_billing_df[clean_matched_idx].iloc[0].to_dict()
    return operator_info


def get_invoice_path(invoice_id: str, _data: dict[str, str]) -> str:
    for k, v in _data.items():
        if invoice_id in k:
            return v
    return ""


def save_invoice_by_number(df: pd.DataFrame, sep: int, save_path: str, path_data: dict):
    dfs = []

    # 通过循环分割DataFrame
    for start in range(0, len(df), sep):
        end = start + sep
        # 使用iloc选取从start到end的行
        split_df = df.iloc[start:end]
        dfs.append(split_df)

    os.makedirs(save_path, exist_ok=True)
    for idx, split_df in enumerate(dfs):
        idx_dir_path = os.path.join(save_path, str(idx))
        os.makedirs(idx_dir_path, exist_ok=True)
        appendix_dir_path = os.path.join(idx_dir_path, "附件")
        os.makedirs(appendix_dir_path, exist_ok=True)
        split_df.to_excel(os.path.join(idx_dir_path, f"service_template.xlsx"), index=False)
        for index, row in split_df.iterrows():
            invoice_id = row["发票号码"]
            invoice_path = get_invoice_path(invoice_id, path_data)
            shutil.copy(invoice_path, os.path.join(appendix_dir_path, f"{invoice_id}.pdf"))
