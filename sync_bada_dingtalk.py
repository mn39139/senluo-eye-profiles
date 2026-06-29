#!/usr/bin/env python3
"""同步八达服饰(福建)有限公司信息到钉钉AI表格（客户资讯表）"""

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

    result = list_records(token)
    records = result.get('value', [])
    print(f"当前记录数: {len(records)}")

    existing_id = None
    for r in records:
        name = r.get('values', {}).get('企业名称', '')
        if name and '八达服饰' in str(name):
            existing_id = r['id']
            print(f"已存在: {name} (ID: {existing_id})")
            break

    fields = {
        "企业名称": "八达服饰（福建）有限公司",
        "统一信用代码": "913505007173546744",
        "法人": "陈其俊",
        "注册资本": "631.15",
        "成立时间": "2000-08-16",
        "注册地址": "福建省泉州经济技术开发区雅泰路92号",
        "年产值": "始建1988年/注册2000年; 外商独资企业; 生产基地20000㎡(绿化40%); 针织梭织面料生产男装/女装/童装/T恤/跑步系列/瑜伽系列/足球系列; 出口欧美为多个世界知名品牌OEM/ODM; 新能源转型(200kw光伏+142KW二期光伏+充电站); 4次新能源招投标; 24条财产线索/预估33万元; 私营不公开披露营业额",
        "营业额": "注册资本631.15万元(实缴631.15万); 2024年税务信用A级; 水滴信用828分; 启信分656分; 参保63人/BOSS直聘100-499人; 薪酬2K-20K/61.3%拿4.5-10K; 8个热招岗位; 13份年报; 36条变更; 4次招投标; 15件司法案件; 1条被执行(8968元)",
        "行业": "运动休闲针织服装制造(C1821)·童装制造·外贸出口OEM/ODM",
        "信用评分": "828",
        "经营范围": "服装制造；服饰制造；服装服饰批发/零售；服装辅料销售；针纺织品销售；日用口罩生产/销售；特种劳动防护用品生产/销售；货物/技术进出口；太阳能发电技术服务；信息系统运行维护服务；电动汽车充电基础设施运营；集中式快速充电站；机动车充电销售；停车场服务",
        "企业画像链接": {"link": "https://mn39139.github.io/senluo-eye-profiles/bada-profile.html"},
        "来源": "森罗之眼智能体",
        "品牌矩阵（产品体系）": "7大产品:男装/女装/童装/T恤/跑步系列/瑜伽系列/足球系列; 面料:针织+梭织; 模式:OEM/ODM为世界知名品牌定牌生产; 出口:欧美地区为主遍布全球; 新能源:200kw光伏一期+142KW二期分布式光伏发电+三期/七期新能源电车充电基础设施; 品牌名:八达",
        "员工规模": "63人(2025年参保/小型企业); BOSS直聘显示100-499人; 薪酬2K-20K(61.3%拿4.5-10K/2025年+5%); 8个热招岗位(亚马逊运营/QC/业务员/外贸业务/跟单/打版/样衣等); 法人陈其俊(法定代表人/100%控股/外商独资/疑似实控人/关联4家); 邱跃飞(主要人员); 本科工资11.1K; 学历要求大专26.5%/高中26.5%/不限24.5%",
        "荣誉资质": "2024年税务信用A级; 一般纳税人; 水滴信用828分(优秀); 启信分656分; 外商独资企业; 35年发展历程(1988-2026); 2项行政许可(污水排入排水管网/道路挖掘); 4次新能源招投标; 24条财产线索; 自身风险6条/司法案件15件/被执行1条(8968元)",
        "联系方式": "电话:0595-28153339 | 邮箱:1847088148@qq.com | 地址:福建省泉州经济技术开发区雅泰路92号 | 发票抬头:福建八达服饰 | 企业族群:福建八达服饰(2家成员)",
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
