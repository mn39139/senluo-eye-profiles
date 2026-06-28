#!/usr/bin/env python3
"""同步奥嘉索集团有限公司信息到钉钉AI表格（客户资讯表）"""

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
        if name and '奥嘉索' in str(name):
            existing_id = r['id']
            print(f"已存在: {name} (ID: {existing_id})")
            break

    # 2. 填充字段
    fields = {
        "企业名称": "奥嘉索集团有限公司",
        "统一信用代码": "91350581MA8RTDYX6P",
        "法人": "洪瑞强",
        "注册资本": "10000",
        "成立时间": "2021-03-30",
        "注册地址": "福建省石狮市伟业路108号奥嘉索工业园",
        "年产值": "峰值年营业额10.6亿元(2014财年12.99亿港元); 2015年10.91亿港元; 2016年8.41亿港元; 2006年5.3亿港元; 巅峰期180+城市1500家店; 财产线索172条估值1.30亿元; 8家子公司; 50件商标",
        "营业额": "峰值10.6亿元人民币(2014财年12.99亿港元); 2015财年10.91亿港元; 2016财年8.41亿港元; 2006财年5.3亿港元; 现转型电商运营天猫Walker Shop旗舰店; 注册资本1亿元; 集团覆盖制鞋/品牌管理/电子商务/供应链/酒店",
        "行业": "皮革、毛皮、羽毛及其制品和制鞋业·皮鞋制造(C1952)·鞋服品牌管理",
        "信用评分": "821",
        "经营范围": "鞋制造; 服装制造; 工程和技术研究和试验发展; 鞋帽批发零售; 服装服饰批发零售; 箱包销售; 皮革制品销售; 服装辅料销售; 针纺织品销售; 互联网销售; 进出口代理; 货物进出口; 技术进出口",
        "企业画像链接": {"link": "https://mn39139.github.io/senluo-eye-profiles/aojiasu-profile.html"},
        "来源": "森罗之眼智能体",
        "品牌矩阵（产品体系）": "WALKER SHOP(核心品牌/1993年香港创立/2007年港交所上市HK1386); COUBER.G(奥巴斯); ARTEMIS(艾迪米斯/商务); ACUPUNCTURE(爱克佩特/英国品牌); Forleria(法娜妮); OX-X-OX(奥乐斯); A+A2(艾迪艾加); MY WALKER; WALKER ONE; Trunari(达斯弥); 50件商标; 2项外观专利(蜗牛鞋等)",
        "员工规模": "100-499人(BOSS直聘/核心制造主体奥卡索福建鞋业); 普利达斯福建鞋业100-499人; 集团参保0人(2024年报/小微企业); 含电商运营/设计/生产/客服团队; 核心管理:洪瑞强(法定代表人/董事/财务负责人/100%控股/疑似实控人/关联企业19家/担任16家法人/15家控股/23家最终受益人); 商业合作伙伴:冯志松/施金田/李春煌/龚玉梅",
        "荣誉资质": "天眼查评分88; 水滴信用821分(优秀); 启信分611; 一般纳税人; 进出口经营权; 2007年港交所主板上市(HK1386); 巅峰期180+城市1300家店铺",
        "联系方式": "电话:0595-88708999 | 地址:福建省石狮市伟业路108号奥嘉索工业园 | 集团成员企业16家",
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
