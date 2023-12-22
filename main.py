from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtCore import pyqtSignal, pyqtSlot, QObject
from PyQt5 import uic
from geocode import get_location
from directions5 import get_optimal_route
import os
import pandas as pd
from openpyxl import load_workbook

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
        worker_df = pd.read_excel(work_file, sheet_name="직원정보")

        # 직원정보 --> 지오코딩
        for i in range(5):
            address = worker_df.loc[i, "실거주주소"]
            name = worker_df.loc[i, "성명"]
            lon = worker_df.loc[i, "LON"]
            if not lon is None:  
                if lon.replace(".", "").isdigit():
                    continue
            lon, lat = get_location(address)
            worker_df.loc[i, "LON"] = lon
            worker_df.loc[i, "LAT"] = lat
            self.signals.progress_changed.emit(i+1, 5)
            if i % 10 == (len(worker_df) % 10 - 1):
                self.signals.messege_added.emit(f'직원정보: {name}') 

        # 직원정보 업데이트
        self.signals.messege_added.emit(f'{work_file} 직원정보 시트를 업데이트합니다')
        if xl_mode == 'w':
            with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode=xl_mode)  as writer:            
                worker_df.to_excel(writer, sheet_name="직원정보", index=False, header=True)
        else:
            with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode=xl_mode, if_sheet_exists='replace')  as writer:            
                worker_df.to_excel(writer, sheet_name="직원정보", index=False, header=True)

        self.signals.messege_added.emit('지사정보 읽는 중입니다')
        jisa_df = pd.read_excel(work_file, sheet_name="지사정보")
        
        # 지사정보 --> 지오코딩
        for i in range(len(jisa_df)):
            address = jisa_df.loc[i, "주소"]
            name = jisa_df.loc[i, "지사명"]
            lon = worker_df.loc[i, "LON"]
            if not lon is None:  
                if lon.replace(".", "").isdigit():
                    continue
            lon, lat = get_location(address)
            jisa_df.loc[i, "LON"] = lon
            jisa_df.loc[i, "LAT"] = lat
            self.signals.progress_changed.emit(i+1, len(jisa_df))
            if i % 10 ==  (len(jisa_df) % 10 - 1):
                self.signals.messege_added.emit(f'지사정보: {name}') 

        # 지사정보 업데이트
        self.signals.messege_added.emit(f'{work_file} 지사정보 시트를 업데이트합니다')
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
  
            self.signals.progress_changed.emit(i+1, len(worker_df))

        # 매칭테이블 업데이트
        self.signals.messege_added.emit(f'{work_file} 출퇴근거리 시트를 업데이트합니다')
        with pd.ExcelWriter(self.work_geocode, engine='openpyxl', mode="a", if_sheet_exists='replace')  as writer:            
            matching_df.to_excel(writer, sheet_name="출퇴근거리", index=False, header=True)

    def make_match_table(self):
        print('matching table')

    def survey_dt(self):
        print('survey dt')

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
        if 0 <_iter <100:
            self.progressBar.setVisible(True)
        else:
            self.progressBar.setVisible(False)
        
        prg = int(_iter / _total * 100)
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
