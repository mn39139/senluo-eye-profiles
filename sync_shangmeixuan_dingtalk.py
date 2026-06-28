#!/usr/bin/env python3
"""同步福建省尚美轩餐饮管理有限公司信息到钉钉AI表格（客户资讯表）"""

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
        if name and '尚美轩' in str(name):
            existing_id = r['id']
            print(f"已存在: {name} (ID: {existing_id})")
            break

    fields = {
        "企业名称": "福建省尚美轩餐饮管理有限公司",
        "统一信用代码": "91350581335741943G",
        "法人": "陈呈宏",
        "注册资本": "1000",
        "成立时间": "2015-04-30",
        "注册地址": "福建省泉州市晋江市永和镇钱仓村诚信路3-1号",
        "年产值": "连锁加盟咖啡&茶饮企业; 始创2014年; 8大品牌矩阵(黑桃漫生活Café/空降茶咖/佬港园/遇福记/茶逅/FIND TEA/爱上漫生活/清心漫生活); 黑桃漫生活品牌百余家直营+加盟店; 13大产品系列(意式咖啡/招牌玛奇朵/鲜果茶/奶茶/甜品等); 130条注册商标; 集品牌推广运营新品研发为一体; 私营不公开披露营业额",
        "营业额": "注册资本1000万元(认缴/2030年到期); 微型企业(参保2人); 年营业额未公开(年报选择不公示); 黑桃漫生活品牌百余家门店(直营+加盟); 顺企网历史显示100人; 0自身风险/0关联风险; 水滴信用820分; 11份企业年报; 1次抽查检查合格",
        "行业": "住宿和餐饮业·连锁加盟咖啡&茶饮·轻食文化",
        "信用评分": "820",
        "经营范围": "餐饮管理；餐饮服务；销售：餐具、包袋。（依法须经批准的项目，经相关部门批准后方可开展经营活动）",
        "企业画像链接": {"link": "https://mn39139.github.io/senluo-eye-profiles/shangmeixuan-profile.html"},
        "来源": "森罗之眼智能体",
        "品牌矩阵（产品体系）": "黑桃漫生活Café(主打品牌/百余家门店); 空降茶咖; 佬港园; 遇福记; 茶逅(注册号23040859/43类); FIND TEA(2021申请/29-32-35-43类); 爱上漫生活(43类); 清心漫生活(43类); 13大产品系列:意式咖啡/现萃茶/招牌玛奇朵/招牌咸奶缇/柠檬鲜果茶/酷乐沙冰/夏季畅饮/醇香奶茶/鲜茶/鲜榨果汁/冬季热饮/手工甜品/现炸小吃; 原料:高海拔红茶/茉香绿茶/乌龙茶",
        "员工规模": "2人(2025年参保/微型企业); 顺企网历史显示100人; 企查查显示少于50人; 法人陈呈宏(执行董事/经理/49%控股/2019年变更); 王志俭(监事/51%控股); 王天沙(原法人/2019年退出); 六大加盟优势; 百余家直营+加盟连锁店",
        "荣誉资质": "水滴信用820分(优秀); 科创分25分; 130条注册商标(8大品牌); 备案网站1个(闽ICP备19016135号); 0自身风险/0关联风险/0司法案件/0行政处罚/0经营异常; 1次抽查检查合格(2016-06-02); 小规模纳税人; 11份企业年报",
        "联系方式": "电话:15959552227/13959962008 | 地址:福建省泉州市晋江市永和镇钱仓村诚信路3-1号 | 官网:www.fjsmxcy.com | 备案:闽ICP备19016135号",
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
