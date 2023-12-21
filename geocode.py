from api_key import AccessID, SecretKey

# *-- Geocoding 활용 코드 --*
import json
import urllib
from urllib.request import Request, urlopen


# 주소에 geocoding 적용하는 함수를 작성.
def get_location(loc) :
    client_id = AccessID
    client_secret = SecretKey
    url = f"https://naveropenapi.apigw.ntruss.com/map-geocode/v2/geocode?query=" + urllib.parse.quote(loc)
    
    # 주소 변환
    request = urllib.request.Request(url)
    request.add_header('X-NCP-APIGW-API-KEY-ID', client_id)
    request.add_header('X-NCP-APIGW-API-KEY', client_secret)
    
    response = urlopen(request)
    res = response.getcode()
    
    if (res == 200) : # 응답이 정상적으로 완료되면 200을 return한다
        response_body = response.read().decode('utf-8')
        response_body = json.loads(response_body)
        # print(response_body)
        # 주소가 존재할 경우 total count == 1이 반환됨.
        if response_body['meta']['totalCount'] > 0 : 
        	# 위도, 경도 좌표를 받아와서 return해 줌.
            lat = response_body['addresses'][0]['y']
            lon = response_body['addresses'][0]['x']
            return (lon, lat)
        else :
            return None, None
    else :
        return None, None
        
if __name__ == '__main__':

    # *-- 3개의 주소 geocoding으로 변환한다.(출발지, 도착지, 경유지) --*
    start = '서울특별시 종로구 청와대로 1'
    goal = '부산광역시 금정구 부산대학로63번길 2'
    waypoint = '경기도 수원시 장안구 서부로 2149'
    #  함수 적용
    start = get_location(start)
    # goal = get_location(goal)
    # waypoint = get_location(waypoint)

    print(start)