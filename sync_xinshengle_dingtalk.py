#!/usr/bin/env python3
"""同步泉州市鑫胜乐商贸有限公司信息到钉钉AI表格（客户资讯表）"""

import requests
import json

APP_KEY = "dingd0hqxztkox98pt12"
APP_SECRET = "FiGX1pghZ2LjodoY8U1QdLVv9B4WeM7ksiInkjQHDurBqfndtfgL9xkPM4pz64uz"
BASE_ID = "o14dA3GK8gard3qvIK1YKY46V9ekBD76"
SHEET_NAME = "客户资讯"
OPERATOR_ID = "TSUgLE1qlXPhYTBrE5FcgwiEiE"


def get_token():
    url = "https://api.dingtalk.com/v1.0/oauth2/accessToken"
    resp = requests.post(url, json={"appKey": APP_KEY, "appSecret": APP_SECRET})
    return resp.json()["accessToken"]


def list_records(token):
    url = f"https://api.dingtalk.com/v1.0/notable/bases/{BASE_ID}/sheets/{SHEET_NAME}/records"
    headers = {"Content-Type": "application/json", "x-acs-dingtalk-access-token": token}
    params = {"operatorId": OPERATOR_ID, "maxResults": 100}
    resp = requests.get(url, headers=headers, params=params)
    return resp.json()


def main():
    token = get_token()
    print(f"Token: {token[:20]}...")

    # 1. 查重
    result = list_records(token)
    records = result.get('value', [])
    print(f"当前记录数: {len(records)}")

    existing_id = None
    for r in records:
        name = r.get('values', {}).get('企业名称', '')
        if name and '鑫胜乐' in str(name):
            existing_id = r['id']
            print(f"已存在: {name} (ID: {existing_id})")
            break

    # 2. 填充字段
    fields = {
        "企业名称": "泉州市鑫胜乐商贸有限公司",
        "统一信用代码": "91350582MA2Y7DGP1M",
        "法人": "谢昌达",
        "注册资本": "50",
        "成立时间": "2017-05-04",
        "注册地址": "福建省泉州市晋江市梅岭街道塘岸街桂华苑18幢2605",
        "年产值": "化工添加剂贸易商; 增韧剂/相容剂/阻燃剂/功能助剂领先供应商; 产品应用于PC/PBT/PET/ABS/PA/PMMA/PP/PE等热塑性材料; 覆盖建筑/电子电器/消费品/运输业; 微型企业(参保3人); 港澳台自然人独资; 2025年税务信用A级",
        "营业额": "注册资本50万元(实缴25万元); 微型企业; 私营不公开披露; 7.5小时工作制双休; 基本工资+提成+绩效奖+年终奖; 进出口贸易; 一般纳税人; 税务A级",
        "行业": "批发业(F5100)·化工添加剂贸易·聚合物解决方案供应商",
        "信用评分": "825",
        "经营范围": "批发零售：五金产品、建材（不含危险化学品）、家具、电子产品、化工产品（不含危险化学品）（以上商品进出口不涉及国营贸易、进出口配额许可证、出口配额招标、出口许可证等专项管理的商品）；普通货物道路运输",
        "企业画像链接": {"link": "https://mn39139.github.io/senluo-eye-profiles/xinshengle-profile.html"},
        "来源": "森罗之眼智能体",
        "品牌矩阵（产品体系）": "增韧剂(改善热塑性聚合物物理特性); 相容剂; 阻燃剂; 功能助剂; 产品标签:涂料/助剂/改性/TPE弹性体/POE/塑胶原料/有机硅/偶联剂/增韧/水性树脂; 应用于PC/PBT/PET/ABS/PA/PMMA/PP/PE; 覆盖建筑/电子电器/消费品/运输业",
        "员工规模": "3人(2025年参保/微型企业); 招聘平台显示20-99人; 核心管理:谢昌达(法定代表人/执行董事/100%控股/港澳台自然人/疑似实控人); 陈如如(财务负责人/2025-06新增); 林宝珍(曾为主要人员); 7.5小时工作制周末双休; 提供住宿含空调/WIFI/热水器/洗衣机",
        "荣誉资质": "2025年税务信用A级; 增值税一般纳税人; 进出口经营权; 港澳台自然人独资企业; 水滴信用825分(优秀); 启信分508分; 0自身风险/0关联风险/0司法案件/0行政处罚",
        "联系方式": "电话:13276009813 | 地址:福建省泉州市晋江市梅岭街道塘岸街桂华苑18幢2605(晋江青阳万达广场旁) | 邮箱:dav***om",
    }

    url = f"https://api.dingtalk.com/v1.0/notable/bases/{BASE_ID}/sheets/{SHEET_NAME}/records"
    headers = {"Content-Type": "application/json", "x-acs-dingtalk-access-token": token}
    params = {"operatorId": OPERATOR_ID}

    if existing_id:
        data = {"records": [{"id": existing_id, "fields": fields}]}
        resp = requests.put(url, headers=headers, params=params, json=data)
        result = resp.json()
        print(f"\n更新结果: {json.dumps(result, ensure_ascii=False)[:500]}")
    else:
        data = {"records": [{"fields": fields}]}
        resp = requests.post(url, headers=headers, params=params, json=data)
        result = resp.json()
        print(f"\n新增结果: {json.dumps(result, ensure_ascii=False)[:500]}")

    print("\nDone")


if __name__ == "__main__":
    main()
