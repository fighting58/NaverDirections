import pandas as pd
from geocode import get_location

jisa = pd.read_excel('거리시간.xlsx', sheet_name="지사")

for i in range(len(jisa)):
    lon, lat = get_location(jisa.loc[i, "주소"])
    jisa.loc[i, "Lon"] = lon
    jisa.loc[i, "Lat"] = lat
    
print(jisa)
