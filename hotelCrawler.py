import requests
import json
import time
import urllib
from openpyxl import Workbook


def get_hotel_info(q, start_day, end_day, off_set):
    # URL地址
    url = "https://ihotel.meituan.com/hbsearch/HotelSearch"
    # 字符转换
    q_urlencoded = urllib.parse.quote(q, safe='/', encoding=None, errors=None)
    start_day_urlencoded = start_day[0:4] + "-" + start_day[4:6] + "-" + start_day[6:8]
    end_day_urlencoded = end_day[0:4] + "-" + end_day[4:6] + "-" + end_day[6:8]
    start_end_day = start_day + "~" + end_day
    # 请求头
    query_headers = {
        "__skcy": "no-signature",
        "Content-Type": "application/json; charset=utf-8",
        "Origin": "https://i.meituan.com",
        "Referer": "https://i.meituan.com/awp/h5/hotel/list/list.html?cityId=55&checkIn=" + start_day_urlencoded + "&checkOut=" + end_day_urlencoded + "&lat=32.351488&lng=119.078614&keyword=" + q_urlencoded + "&accommodationType=1&sort=smart&ste=_b400203",
        "Sec-Fetch-Mode": "cors",
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Mobile Safari/537.36"
    }
    # 请求体
    query_params = {
        "utm_medium": "touch",
        "version_name": "999.9",
        "platformid": "1",
        "cateId": "20",
        "newcate": "1",
        "limit": "20",
        "offset": str(off_set),
        "cityId": "55",
        "ci": "55",
        "startendday": start_end_day,
        "startDay": start_day,
        "endDay": start_day,
        "q": q,
        "ste": "_b400203",
        "mypos": "32.351488,119.078614",
        "attr_28": "129",
        "sort": "rating",
        "price": "0~999999",
        "uuid": "C87272C89D612AD08DBDAC8A29A2DFDF25AD73E091CEE41C1371DA8304BD24BF"
    }
    # 发起、处理请求
    response = requests.get(url, query_params, headers=query_headers)
    result = json.loads(response.text)
    total_count = result["data"]["totalcount"]
    hotel_info = result["data"]["searchresult"]
    print("本次搜索共" + str(total_count) + "条记录:")
    for hotel in hotel_info:
        print("酒店名称：" + hotel["name"])
        print("所在城市：" + hotel["cityName"])
        print("酒店地址：" + hotel["addr"])
        print("酒店品阶：" + hotel["hotelStar"])
        print("酒店最低价：" + str(hotel["lowestPrice"]))
        print("酒店评分：" + str(hotel["avgScore"]))
        print("距离搜索目标位置：" + hotel["posdescr"])
    return total_count, hotel_info


def contain_any(strings, string):
    for an in strings:
        if an in string:
            return True
    return False


def save_hotel_info(q, start_day, end_day):
    # 写入本地excel文件中
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = u"hotelInfo"
        ws.append(["酒店名称", "所在城市", "酒店地址", "酒店品阶", "酒店最低价", "酒店评分", "距离搜索目标位置"])
        total_count, _ = get_hotel_info(q, start_day, end_day, 0)
        i = 0
        hotel_data = []
        while i < total_count:
            _, data = get_hotel_info(q, start_day, end_day, i)
            i += 20
            hotel_data += data
            time.sleep(10)
        for hotel in hotel_data:
            if contain_any(["青旅", "青年", "旅馆", "旅舍", "客栈", "宾馆"], hotel["name"]):
                continue
            else:
                data = ["", "", "", "", "", "", ""]
                data[0] = hotel["name"]
                data[1] = hotel["cityName"]
                data[2] = hotel["addr"]
                data[3] = hotel["hotelStar"]
                data[4] = hotel["lowestPrice"]
                data[5] = hotel["avgScore"]
                data[6] = hotel["posdescr"]
                ws.append(data)
        wb.save("hotelInfo.xlsx")
    except Exception as e:
        print("Error：", e)
    finally:
        print("程序运行结束！")


# 主函数
if __name__ == '__main__':
    save_hotel_info("天隆寺", "20191010", "20191011")
