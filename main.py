from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtCore import pyqtSignal, pyqtSlot, QObject
from PyQt5 import uic
from geocode import get_location
from directions5 import get_optimal_route
import os
import pandas as pd
from openpyxl import load_workbook
import numpy as np
import time

form_class = uic.loadUiType('UI\main.ui')[0]


class CustomSignals(QObject):
    messege_added = pyqtSignal(str)
    progress_changed = pyqtSignal(int, int)


class Ui_Form(QWidget, form_class):

    signals = CustomSignals()
    work_file = "직원명부.xlsx"
    work_geocode = "직원명부_geocode.xlsx"

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.butSurveyGeocode.clicked.connect(self.do_geocoding)
        self.butSurveyDT.clicked.connect(self.survey_dt)
        self.txtSearchCoordinate.clicked.connect(self.single_geocoding)
        self.butSearchDistance.clicked.connect(self.single_dt)

        self.signals.progress_changed.connect(self.progrss)
        self.signals.messege_added.connect(self.show_message)
        self.progressBar.setVisible(False)

    def do_geocoding(self):
        
        work_file = self.work_file
        xl_mode = 'w'
        # 직원명부_geocode가 작성되어 있으면 그것을 대상으로 작업
        if os.path.exists(self.work_geocode):
            xl_mode = "a"
            work_file = self.work_geocode
        
        # Dataframe으로 직원 및 지사정보를 가져옴
        self.signals.messege_added.emit('직원정보를 읽는 중입니다')
        worker_df = pd.read_excel(work_file, sheet_name="직원정보", dtype={"사번": str, "LON":str, "LAT":str})
        worker_df.reindex()

        # 직원정보 --> 지오코딩
        for i in range(len(worker_df)):
            address = worker_df.loc[i, "실거주주소"]
            lon = worker_df.loc[i, "LON"]
            if not lon is None:  
                if str(lon).replace(".", "").isdigit():
                    continue
            lon, lat = get_location(address)
            worker_df.loc[i, "LON"] = lon
            worker_df.loc[i, "LAT"] = lat
            self.signals.progress_changed.emit(i+1, len(worker_df))

        # 직원정보 업데이트
        self.signals.messege_added.emit(f'{self.work_geocode} 직원정보 시트를 업데이트합니다')
        if xl_mode == 'w':
            with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode=xl_mode)  as writer:            
                worker_df.to_excel(writer, sheet_name="직원정보", index=False, header=True)
        else:
            with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode=xl_mode, if_sheet_exists='replace')  as writer:            
                worker_df.to_excel(writer, sheet_name="직원정보", index=False, header=True)

        self.signals.messege_added.emit('지사정보 읽는 중입니다')
        jisa_df = pd.read_excel(work_file, sheet_name="지사정보", dtype={"LON":str, "LAT":str})
        jisa_df.reindex()
        
        # 지사정보 --> 지오코딩
        for i in range(len(jisa_df)):
            address = jisa_df.loc[i, "주소"]
            lon = jisa_df.loc[i, "LON"]
            if not lon is None:  
                if str(lon).replace(".", "").isdigit():
                    continue
            lon, lat = get_location(address)
            jisa_df.loc[i, "LON"] = lon
            jisa_df.loc[i, "LAT"] = lat
            self.signals.progress_changed.emit(i+1, len(jisa_df))

        # 지사정보 업데이트
        self.signals.messege_added.emit(f'{self.work_geocode} 지사정보 시트를 업데이트합니다')
        with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode="a", if_sheet_exists='replace')  as writer:            
            jisa_df.to_excel(writer, sheet_name="지사정보", index=False, header=True)

        # 매칭데이블 작성
        self.signals.messege_added.emit('매칭 테이블을 작성중입니다')
        matching_df = pd.DataFrame(columns=["사번", "성명", "LON", "LAT", "지사명", "지사_LON", "지사_LAT", "거리", "시간"], dtype=str)
        for i in range(len(worker_df)):
            worker_info = [v for v in worker_df.loc[i, ["사번", "성명", "LON", "LAT"]].values]
            for j in range(len(jisa_df)):
                jisa_info = [v for v in jisa_df.loc[j, ["지사명", "LON", "LAT"]].values]            
                match_list = worker_info + jisa_info + [None, None]
                new_df = pd.DataFrame([match_list], columns=matching_df.columns)
                matching_df= pd.concat([matching_df, new_df])
            if i % 50 == 0:
                self.signals.progress_changed.emit(i+1, len(worker_df))
        self.signals.progress_changed.emit(i+1, len(worker_df))

        # 매칭테이블 업데이트
        self.signals.messege_added.emit(f'{self.work_geocode} 출퇴근거리 시트를 업데이트합니다')
        with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode="a", if_sheet_exists='replace')  as writer:            
            matching_df.to_excel(writer, sheet_name="출퇴근거리", index=False, header=True)
        
        self.signals.messege_added.emit(f'{self.work_geocode}: 지오코딩 및 매칭테이블 작성이 완료되었습니다.')

    def survey_dt(self):
       
        # 매칭테이블을 데이터프레임으로 가져옴
        
        self.signals.messege_added.emit('매칭 테이블을 가져옵니다')
        matching_df = pd.read_excel(self.work_geocode, sheet_name="출퇴근거리", dtype=str)
        matching_df.reindex()

        # 자택-지사간 거리/시간 조회
        self.signals.messege_added.emit('자택-지사간 거리/시간을 조회합니다')
        start_time = time.time()
        for i in range(len(matching_df)):

            # 기 조회된 값이 있으면 패스(기록된 값이 숫자로 된 값이 아닐 경우)
            duration = matching_df.loc[i, "시간"]
            name = matching_df.loc[i, "성명"]

            if not pd.isna(np.array(duration)):  
                if str(duration).isdigit():
                    continue

            # 각 좌표값에 오류가 있을 경우 패스
            coords = [v for v in matching_df.loc[i, ["LON", "LAT", "지사_LON", "지사_LAT"]]]
            if ("N_A" in coords) or ("COM" in coords):
                continue

            start = [v for v in matching_df.loc[i, ["LON", "LAT"]].values]  # 자택 좌표
            goal = [v for v in matching_df.loc[i, ["지사_LON", "지사_LAT"]].values]  #  지사 좌표
            d_t = get_optimal_route(start, goal)  # 거리/시간 조회
            matching_df.loc[i, "거리"] = d_t["total_distance"]
            matching_df.loc[i, "시간"] = d_t["total_duration"]

            self.signals.progress_changed.emit(i+1, len(matching_df))
        print(f'{time.time()-start_time}')

        self.signals.messege_added.emit('매칭 테이블을 업데이트합니다')
        with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode="a", if_sheet_exists='replace')  as writer:            
            matching_df.to_excel(writer, sheet_name="출퇴근거리", index=False, header=True)
        
        self.signals.messege_added.emit(f'{self.work_geocode}, 작성이 완료되었습니다.')


    def single_geocoding(self):
        address = self.txtSingleAddress.text()
        lon, lat = get_location(address)
        self.txtLon.setText(lon)
        self.txtLat.setText(lat)

    def single_dt(self):
        start_address = self.txtStartAddress.text()
        goal_address = self.txtDepatureAddress.text()
        start = get_location(start_address)
        goal = get_location(goal_address)
        d_t = get_optimal_route(start, goal)
        self.txtDistance.setText(d_t["total_distance"])
        self.txtTime.setText(d_t["total_duration"])

    @pyqtSlot(int, int)
    def progrss(self, _iter, _total):
        prg = int(_iter / _total * 100)
        if 1 < prg < 99:
            self.progressBar.setVisible(True)
        else:
            self.progressBar.setVisible(False)
        self.progressBar.setValue(prg)

    @pyqtSlot(str)
    def show_message(self, msg):
        self.lblMessage.setText(msg)


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    ui = Ui_Form()
    ui.show()
    sys.exit(app.exec_())
