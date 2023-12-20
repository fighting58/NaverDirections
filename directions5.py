from api_key import AccessID, SecretKey

import json
import urllib
from urllib.request import Request, urlopen
from geocode import get_location

# *-- Directions 5 활용 코드 --*

def waypoints2string(wpoints:list) -> str:
    val = ""
    if wpoints and isinstance(wpoints, list):
        for i, coord in enumerate(wpoints):
            val += str(coord)
            if i % 2 == 1:
                val += ":"
        return val.rstrip(":")
    elif not isinstance(wpoints, list):
        print("Not List: Waypoints must be an instance of List")
        return 
    elif not wpoints:
        return
    else:
        print("Unknown Error: waypoints2string")
        return

def miliseconds_to_hms(miliseconds: int) -> str:
    # 시간, 분, 초를 계산합니다.
    sec = miliseconds / 1000
    hours = sec // 3600
    minutes = (sec % 3600) // 60
    seconds = sec % 60

    # 시간, 분, 초를 hh:mm:ss 형식으로 포맷팅합니다.
    time_str = "{:02}:{:02}:{:02}".format(int(hours), int(minutes), int(seconds))

    return time_str

def miliseconds_to_minutes(miliseconds: int) -> str:
    sec = miliseconds / 1000
    hours = sec // 3600
    minutes = (sec % 3600) // 60

    return str(int(hours*60 + minutes))

def meter2kilometer(meter: int) -> str:
    return str(meter/1000)
        
def get_optimal_route(start_pos: list, goal_pos: list, waypoints: list = [], option: str ='traoptimal') -> dict:
    # waypoint는 최대 5개까지 입력 가능, 
    # 구분자로 |(pipe char) 사용하면 됨(x,y 좌표값으로 넣을 것)
    # waypoint 옵션을 다수 사용할 경우, 아래 함수 포맷을 바꿔서 사용 
    # option : 탐색옵션 [최대 3개, 여러 옵션은 ':'로 연결, traoptimal(기본 옵션:실시간 최적) 
    # / trafast:실시간 빠른길, tracomfort:실시간 편한길, traavoidtoll:무료우선, traavoidcaronly:자동자 전용도로 회피 우선]
    # start=/goal=/(waypoint=)/(option=) 순으로 request parameter 지정
    waypoints_str = waypoints2string(waypoints)
    error_message = ""

    if not start_pos:
        error_message += "START, "
    if not goal_pos:
        error_message += "GOAL"

    if error_message: 
        return {'total_distance' : f'좌표변환 에러: {error_message.rstrip(", ")}',
                'total_duration' : "Error"}

    else:
        if waypoints_str:
            url = f"https://naveropenapi.apigw.ntruss.com/map-direction/v1/driving?start={start_pos[0]},{start_pos[1]}&goal={goal_pos[0]},{goal_pos[1]}&waypoints={waypoints[0]},{waypoints[1]}&option={option}"
        else:
            url = f"https://naveropenapi.apigw.ntruss.com/map-direction/v1/driving?start={start_pos[0]},{start_pos[1]}&goal={goal_pos[0]},{goal_pos[1]}&option={option}"

        request = urllib.request.Request(url)
        request.add_header('X-NCP-APIGW-API-KEY-ID', AccessID)
        request.add_header('X-NCP-APIGW-API-KEY', SecretKey)
        
        response = urllib.request.urlopen(request)
        res = response.getcode()
        
        if (res == 200) :
            response_body = response.read().decode('utf-8')
            results = json.loads(response_body)
            return {'total_distance' : meter2kilometer(results['route']['traoptimal'][0]['summary']['distance']),
                    'total_duration' : miliseconds_to_minutes(results['route']['traoptimal'][0]['summary']['duration'])}
                
        else:
            return {'total_distance' : "통신 에러: Error",
                    'total_duration' : "Error"}
        
if __name__ == '__main__':
            
    start = '수원시 팔달구 인계로 21'
    goal = '용인시 처인구 중부대로 1414'

    start_pos = get_location(start)
    goal_pos = get_location(goal)

    summary = get_optimal_route(start_pos, goal_pos)
    print(summary)