#!/usr/bin/env python3
"""同步福建省三恒体育科技有限公司信息到钉钉AI表格（客户资讯表）"""

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
        if name and '三恒' in str(name):
            existing_id = r['id']
            print(f"已存在: {name} (ID: {existing_id})")
            break

    # 2. 填充字段
    fields = {
        "企业名称": "福建省三恒体育科技有限公司",
        "统一信用代码": "91350581MA321MQT24",
        "法人": "李金满",
        "注册资本": "500",
        "成立时间": "2018-08-28",
        "注册地址": "福建省石狮市锦尚镇杨厝路西59号",
        "年产值": "体育用品制造与销售; 25个注册商标; 2项外观专利(上衣/运动包); 淘宝三恒品牌篮球裁判服700+人付款; 参与校服/体育用品政府采购招投标; 关联集团三联服饰(3家成员/注册资本1558万/社保24人/39项专利/9次招投标); 2024年减资1000万→500万; 私营不公开披露营业额",
        "营业额": "注册资本500万元(实缴105万元); 2024年减资(原1000万→500万); 微型企业(参保4人); 2025年税务信用A级; 水滴信用832分; 私营不公开披露; 淘宝裁判服单价32-36元/700+人付款; 参与招投标",
        "行业": "研究和试验发展(R7320)·体育用品制造·服装制造·软件开发",
        "信用评分": "832",
        "经营范围": "工程和技术研究和试验发展；体育用品及器材制造/批发/零售；服装制造/服饰制造/服装服饰批发/零售；服装辅料制造/销售；新材料技术推广服务；体育竞赛组织；皮革制品销售；产业用纺织制成品销售；日用品销售；互联网销售；软件开发/销售；数字文化创意软件开发；网络与信息安全软件开发；软件外包服务；专业设计服务；技术服务/开发/咨询/交流/转让/推广；数据处理服务；物联网技术研发；信息技术咨询服务；网络技术服务；信息系统运行维护服务；计算机软硬件及辅助设备零售/批发；信息系统集成服务；云计算设备销售；区块链技术相关软件和服务",
        "企业画像链接": {"link": "https://mn39139.github.io/senluo-eye-profiles/sanheng-profile.html"},
        "来源": "森罗之眼智能体",
        "品牌矩阵（产品体系）": "体育用品制造(运动服装/裁判服/校服); 服装制造(运动服套装/学生校服/裁判服定制); 软件开发(数字文化创意软件/网络与信息安全软件); 体育竞赛组织; 招投标(校服/体育用品政府采购); 25个注册商标(25类服装鞋帽); 2项外观专利(上衣119285/运动包); 淘宝三恒品牌裁判服",
        "员工规模": "4人(2025年参保/微型企业); 法人李金满(法定代表人/执行董事/经理/财务负责人/100%控股); 侯秀兰(监事/关联5家企业); 集团社保24人(2024年); 关联企业:石狮市三联服饰织造(80%控股/1058万注册资本/2002年成立)、深圳市及时选科技(6%持股/监事)",
        "荣誉资质": "2025年税务信用A级; 一般纳税人; 水滴信用832分(优秀); 科创分43分; 25个注册商标; 2项外观专利; 战略新兴产业(数字文化创意软件开发); 1条自身风险(计算机软件开发合同纠纷)",
        "联系方式": "电话:0595-88752786 | 地址:福建省石狮市锦尚镇杨厝路西59号 | 官网:www.sanhengsports.cn | 英文名:Fujian Sanheng Sports Technology Co., Ltd.",
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
