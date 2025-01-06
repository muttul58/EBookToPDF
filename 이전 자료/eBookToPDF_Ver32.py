# 필요한 모듈들을 임포트합니다.
import sys  # 시스템 특정 파라미터와 함수를 제공합니다.
import time  # 시간 관련 함수를 제공합니다.
import pyautogui  # GUI 자동화를 위한 크로스 플랫폼 모듈입니다.
import os  # 운영 체제와 상호 작용하기 위한 함수를 제공합니다.
import io  # 입출력 작업을 위한 핵심 도구를 제공합니다.
import shutil  # 고수준 파일 연산을 제공합니다.
import logging  # 로깅 기능을 제공합니다.
import traceback  # 예외 추적 정보를 제공합니다.
import tempfile  # 임시 파일 및 디렉토리 생성을 위한 모듈입니다.
import pygetwindow as gw  # 창 관리를 위한 모듈입니다.
import subprocess  # 새로운 프로세스를 생성하고 관리합니다.

# PyQt5에서 필요한 클래스들을 임포트합니다.
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QSpinBox, QSlider, QWidget, QMessageBox,
                             QFileDialog, QScrollArea, QTabWidget, QGroupBox, QListWidget,
                             QSizePolicy, QProgressBar, QCheckBox, QButtonGroup, QRadioButton, QRadioButton, QSpacerItem, QSizePolicy)
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QFont, QIcon
from PyQt5.QtCore import Qt, QRect, QPoint, QSize

# ReportLab 라이브러리에서 필요한 클래스들을 임포트합니다.
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.pagesizes import A4, landscape
from PIL import Image  # 이미지 처리를 위한 Pillow 라이브러리를 임포트합니다.
from PyPDF2 import PdfReader, PdfWriter  # PDF 파일 처리를 위한 PyPDF2 라이브러리를 임포트합니다.
from reportlab.lib.utils import ImageReader
from datetime import datetime  # 시스템 날짜, 시간 가져오기 위한 임포트합니다.

# 로깅 설정을 초기화합니다.
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 이미지 위젯 클래스를 정의합니다.
class ImageWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)  # 부모 클래스의 __init__ 메서드를 호출합니다.
        self.parent = parent  # 부모 위젯을 저장합니다.
        self.pixmap = None  # 표시할 이미지를 저장합니다.
        self.rubberband = QRect()  # 선택 영역을 나타내는 사각형입니다.
        self.origin = QPoint()  # 선택 시작점입니다.
        self.current_rect = QRect()  # 현재 이미지 영역입니다.
        self.dragging = False  # 드래그 중인지 여부를 나타냅니다.
        self.zoom_factor = 1.0  # 줌 배율입니다.
        self.offset = QPoint(0, 0)  # 이미지 오프셋입니다.
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 크기 정책을 설정합니다.
        
        

    def setPixmap(self, pixmap):
        self.pixmap = pixmap  # 새 이미지를 설정합니다.
        self.fit_image()  # 이미지를 위젯에 맞게 조정합니다.
        self.update()  # 위젯을 다시 그립니다.

    def fit_image(self):
        if self.pixmap:  # 이미지가 있는 경우
            self.current_rect = self.rect()  # 현재 위젯의 크기로 설정합니다.
            self.zoom_factor = min(self.width() / self.pixmap.width(),
                                   self.height() / self.pixmap.height())  # 줌 팩터를 계산합니다.
            scaled_size = self.pixmap.size() * self.zoom_factor  # 스케일된 크기를 계산합니다.
            self.current_rect.setSize(scaled_size)  # 현재 사각형의 크기를 설정합니다.
            self.current_rect.moveCenter(self.rect().center())  # 사각형을 중앙에 배치합니다.
            self.offset = QPoint(0, 0)  # 오프셋을 초기화합니다.

    def paintEvent(self, event):
        if self.pixmap:  # 이미지가 있는 경우
            painter = QPainter(self)  # QPainter 객체를 생성합니다.
            scaled_pixmap = self.pixmap.scaled(self.current_rect.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)  # 이미지를 스케일링합니다.
            painter.drawPixmap(self.current_rect, scaled_pixmap)  # 이미지를 그립니다.
            if not self.rubberband.isNull():  # 선택 영역이 있는 경우
                painter.setPen(QPen(Qt.red, 2, Qt.SolidLine))  # 펜 설정을 변경합니다.
                painter.drawRect(self.rubberband)  # 선택 영역을 그립니다.

    def sizeHint(self):
        return QSize(1000, 800)  # 기본 크기를 제안합니다.

    def wheelEvent(self, event):
        if self.pixmap:  # 이미지가 있는 경우
            old_pos = self.mapToPixmap(event.pos())  # 현재 마우스 위치를 이미지 좌표로 변환합니다.
            zoom_change = 1.1 if event.angleDelta().y() > 0 else 1/1.1  # 줌 변화량을 계산합니다.
            self.zoom_factor *= zoom_change  # 줌 팩터를 업데이트합니다.
            self.zoom_factor = max(0.1, min(5.0, self.zoom_factor))  # 줌 팩터를 제한합니다.
            new_size = self.pixmap.size() * self.zoom_factor  # 새 크기를 계산합니다.
            self.resize(new_size)  # 위젯 크기를 조정합니다.
            new_pos = self.mapFromPixmap(old_pos)  # 새 위치를 계산합니다.
            self.offset += event.pos() - new_pos  # 오프셋을 업데이트합니다.
            self.update()  # 위젯을 다시 그립니다.
            self.parent.adjust_scroll_bar()  # 스크롤바를 조정합니다.

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:  # 왼쪽 버튼을 클릭한 경우
            self.origin = event.pos()  # 시작점을 설정합니다.
            self.rubberband = QRect()  # 선택 영역을 초기화합니다.
            self.dragging = True  # 드래그 상태를 설정합니다.

    def mouseMoveEvent(self, event):
        if self.dragging:  # 드래그 중인 경우
            self.rubberband = QRect(self.origin, event.pos()).normalized()  # 선택 영역을 업데이트합니다.
            self.rubberband = self.rubberband.intersected(self.current_rect)  # 이미지 영역 내로 제한합니다.
            self.update()  # 위젯을 다시 그립니다.

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.dragging:  # 왼쪽 버튼을 놓고 드래그 중이었던 경우
            self.dragging = False  # 드래그 상태를 해제합니다.
            if not self.rubberband.isNull():  # 선택 영역이 있는 경우
                self.parent.update_crop_coordinates(self.rubberband)  # 크롭 좌표를 업데이트합니다.

    def mapToPixmap(self, pos):
        if self.current_rect.contains(pos):  # 위치가 현재 이미지 영역 내에 있는 경우
            relative_x = (pos.x() - self.current_rect.left()) / self.current_rect.width()  # 상대적 x 좌표를 계산합니다.
            relative_y = (pos.y() - self.current_rect.top()) / self.current_rect.height()  # 상대적 y 좌표를 계산합니다.
            return QPoint(int(relative_x * self.pixmap.width()),
                          int(relative_y * self.pixmap.height()))  # 이미지 좌표로 변환하여 반환합니다.
        return QPoint()  # 영역 밖인 경우 빈 포인트를 반환합니다.

    def mapFromPixmap(self, pos):
        relative_x = pos.x() / self.pixmap.width()  # 상대적 x 좌표를 계산합니다.
        relative_y = pos.y() / self.pixmap.height()  # 상대적 y 좌표를 계산합니다.
        return QPoint(int(relative_x * self.current_rect.width() + self.current_rect.left()),
                      int(relative_y * self.current_rect.height() + self.current_rect.top()))  # 위젯 좌표로 변환하여 반환합니다.

    def fit_image_to_view(self):
        self.fit_to_view = True  # 뷰에 맞추기 플래그를 설정합니다.
        self.zoom_factor = 1.0  # 줌 팩터를 초기화합니다.
        self.offset = QPoint()  # 오프셋을 초기화합니다.
        self.update()  # 위젯을 다시 그립니다.
        self.updateGeometry()  # 위젯의 지오메트리를 업데이트합니다.
        self.parent.adjust_scroll_bar()  # 스크롤바를 조정합니다.

    def resizeEvent(self, event):
        super().resizeEvent(event)  # 부모 클래스의 resizeEvent를 호출합니다.
        if self.pixmap:  # 이미지가 있는 경우
            self.fit_image()  # 이미지를 다시 맞춥니다.
            self.update()  # 위젯을 다시 그립니다.

class IntegratedEBookApp(QMainWindow):
    def __init__(self):
        super().__init__()  # 부모 클래스의 __init__ 메서드를 호출합니다.
        self.setWindowTitle("eBook 캡처 및 PDF 생성기 by muttul")  # 윈도우 제목을 설정합니다.
        self.setGeometry(100, 100, 1000, 1000)  # 윈도우 크기와 위치를 설정합니다. (x, y, width, height)

        # 전역 폰트를 설정합니다.
        font = QFont("돋움", 9)  # '돋움' 폰트, 크기 9로 QFont 객체를 생성합니다.
        QApplication.setFont(font)  # 생성한 폰트를 애플리케이션 전체에 적용합니다.
        
        # 애플리케이션의 스타일을 설정합니다.
        self.setStyleSheet("""
            QMainWindow, QWidget {
                font-family: '돋움';   /* # 모든 QMainWindow와 QWidget에 '돋움' 폰트를 적용합니다. */
                font-size: 9pt;   /* # 폰트 크기를 9포인트로 설정합니다. */
            }
            QLabel {
                color: #333;   /* # QLabel의 텍스트 색상을 어두운 회색(#333)으로 설정합니다. */
            }
            QPushButton {
                background-color: #4CAF50;   /* # 버튼 배경색을 초록색으로 설정합니다. */
                color: white;   /* # 버튼 텍스트 색상을 흰색으로 설정합니다. */
                border-radius: 4px;  /* # 버튼 모서리를 4픽셀 둥글게 만듭니다. */
                padding: 6px;  /*  버튼 내부에 6픽셀의 패딩을 추가합니다. */
                font-weight: bold;  /* # 버튼 텍스트를 굵게 설정합니다. */
            }
            QPushButton:hover {
                background-color: #45a049;  /*  마우스를 올렸을 때 버튼 배경색을 약간 어둡게 변경합니다. */
            }
            QPushButton:pressed {
                background-color: #39843c;  /* # 버튼을 눌렀을 때 배경색을 더 어둡게 변경합니다. */
            }
            QSpinBox, QSlider, QListWidget {
                background-color: #ffffff;  /* # 배경색을 흰색으로 설정합니다. */
                border: 1px solid #bbb;  /* # 1픽셀 두께의 회색 테두리를 추가합니다. */
                padding: 4px;  /* # 내부에 4픽셀의 패딩을 추가합니다. */
            }
            QTabWidget::pane {
                border: 1px solid #d7d7d7;  /* # 탭 위젯 패널에 1픽셀 두께의 밝은 회색 테두리를 추가합니다. */
                background-color: #ffffff;  /* # 탭 위젯 패널의 배경색을 흰색으로 설정합니다. */
            }
            QTabBar::tab {
                background-color: #e1e1e1;  /* # 탭의 배경색을 밝은 회색으로 설정합니다. */
                padding: 8px 12px;  /* # 탭 내부에 상하 8픽셀, 좌우 12픽셀의 패딩을 추가합니다. */
                margin-right: 2px;  /* # 탭 사이에 2픽셀의 간격을 추가합니다. */
                min-width: 100px;  /* # 탭의 최소 너비를 100픽셀로 설정합니다. */
            }
            QTabBar::tab:selected {
                background-color: #ffffff;  /* # 선택된 탭의 배경색을 흰색으로 설정합니다. */
                font-weight: bold;  /* # 선택된 탭의 텍스트를 굵게 설정합니다. */
            }
            QGroupBox {
                font-weight: bold;  /* # 그룹박스 제목을 굵게 설정합니다. */
                border: 2px solid #bbb;  /* # 2픽셀 두께의 회색 테두리를 추가합니다. */
                border-radius: 4px;  /* # 그룹박스 모서리를 4픽셀 둥글게 만듭니다. */
                margin-top: 6px;  /* # 위쪽에 6픽셀의 마진을 추가합니다. */
                padding-top: 6px;  /* # 위쪽에 6픽셀의 패딩을 추가합니다. */
            }
            QGroupBox::title {
                subcontrol-origin: margin;  /* # 제목의 위치를 마진에서부터 시작하도록 설정합니다. */
                left: 8px;  /* # 제목을 왼쪽에서 8픽셀 떨어뜨립니다. */
                padding: 0 3px 0 3px;  /* # 제목 좌우에 3픽셀의 패딩을 추가합니다. */
            }
        """)  # 위의 스타일시트를 애플리케이션에 적용합니다.
        
        # 기본 경로를 설정합니다.
        self.default_path = os.path.expanduser("~\\Videos\\Desktop")  # 사용자의 비디오/데스크톱 폴더를 기본 경로로 설정합니다.
        self.base_folder = self.default_path  # 기본 폴더를 default_path로 설정합니다.
        self.image_folder = os.path.join(self.base_folder, "Image")  # 이미지 폴더 경로를 설정합니다.
        self.cropper_folder = os.path.join(self.base_folder, "Cropper")  # 크로퍼 폴더 경로를 설정합니다.
        
        # 중앙 위젯과 메인 레이아웃을 설정합니다.
        self.central_widget = QWidget()  # 중앙 위젯을 생성합니다.
        self.setCentralWidget(self.central_widget)  # 중앙 위젯을 메인 윈도우에 설정합니다.
        self.main_layout = QVBoxLayout(self.central_widget)  # 중앙 위젯에 수직 박스 레이아웃을 설정합니다.

        # 탭 위젯을 생성하고 메인 레이아웃에 추가합니다.
        self.tab_widget = QTabWidget()  # 탭 위젯을 생성합니다.
        self.main_layout.addWidget(self.tab_widget)  # 탭 위젯을 메인 레이아웃에 추가합니다.

        # 캡처 탭과 PDF 탭을 생성합니다.
        self.capture_tab = QWidget()  # 캡처 탭 위젯을 생성합니다.
        self.capture_layout = QVBoxLayout(self.capture_tab)  # 캡처 탭에 수직 박스 레이아웃을 설정합니다.
        self.tab_widget.addTab(self.capture_tab, "화면 캡처")  # 캡처 탭을 탭 위젯에 추가합니다.

        self.pdf_tab = QWidget()  # PDF 탭 위젯을 생성합니다.
        self.pdf_layout = QVBoxLayout(self.pdf_tab)  # PDF 탭에 수직 박스 레이아웃을 설정합니다.
        self.tab_widget.addTab(self.pdf_tab, "PDF 생성")  # PDF 탭을 탭 위젯에 추가합니다.

        # 각 탭의 UI를 설정합니다.
        self.setup_capture_tab()  # 캡처 탭 UI를 설정하는 메서드를 호출합니다.
        self.setup_pdf_tab()  # PDF 탭 UI를 설정하는 메서드를 호출합니다.

        # 창 제목 리스트와 선택된 창을 초기화합니다.
        self.window_titles = []  # 창 제목 리스트를 빈 리스트로 초기화합니다.
        self.selected_window = None  # 선택된 창을 None으로 초기화합니다.
        
        # 탭 변경 시 이벤트를 연결합니다.
        self.tab_widget.currentChanged.connect(self.on_tab_changed)  # 탭이 변경될 때 on_tab_changed 메서드를 호출합니다.
        
        # PDF 탭 초기화 여부를 나타내는 플래그입니다.
        self.is_pdf_tab_initialized = False  # PDF 탭 초기화 상태를 False로 설정합니다.

    # 캡처 탭 UI를 설정하는 메서드입니다.
    def setup_capture_tab(self):
        self.capture_layout = QVBoxLayout()  # 캡처 탭의 메인 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        self.capture_layout.setSpacing(20)  # 레이아웃 내 위젯 간 간격을 20픽셀로 설정합니다.

        # 캡처 설정 그룹을 생성합니다.
        settings_group = QGroupBox("캡처 설정")  # "캡처 설정"이라는 제목의 그룹 박스를 생성합니다.
        settings_layout = QVBoxLayout()  # 설정 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.

        # 반복 횟수 설정 UI를 생성합니다.
        repeat_layout = QHBoxLayout()  # 반복 횟수 설정을 위한 수평 박스 레이아웃을 생성합니다.
        self.repeat_label = QLabel("반복 횟수 :     ", self)  # "반복 횟수 : " 레이블을 생성합니다.
        repeat_layout.addWidget(self.repeat_label)  # 레이블을 레이아웃에 추가합니다.

        self.repeat_spinbox = QSpinBox(self)  # 반복 횟수를 입력받을 스핀 박스를 생성합니다.
        self.repeat_spinbox.setRange(1, 2000)  # 스핀 박스의 범위를 1부터 2000까지로 설정합니다.
        self.repeat_spinbox.setValue(1)  # 스핀 박스의 초기값을 1로 설정합니다.
        self.repeat_spinbox.setFixedHeight(35)  # 스핀 박스의 높이를 35픽셀로 고정합니다.
        repeat_layout.addWidget(self.repeat_spinbox, 1)  # 스핀 박스를 레이아웃에 추가하고, 늘어나는 공간을 차지하도록 설정합니다.

        settings_layout.addLayout(repeat_layout)  # 반복 횟수 설정 레이아웃을 설정 그룹 레이아웃에 추가합니다.

        # 대기 시간 설정 UI를 생성합니다.
        wait_layout = QHBoxLayout()  # 대기 시간 설정을 위한 수평 박스 레이아웃을 생성합니다.
        self.slider_label = QLabel("대기 시간(초) :", self)  # "대기 시간(초) :" 레이블을 생성합니다.
        wait_layout.addWidget(self.slider_label)  # 레이블을 레이아웃에 추가합니다.
        self.time_slider = QSlider(Qt.Horizontal, self)  # 수평 슬라이더를 생성합니다.
        self.time_slider.setRange(1, 20)  # 슬라이더의 범위를 1부터 20까지로 설정합니다.
        self.time_slider.setValue(1)  # 슬라이더의 초기값을 1로 설정합니다.
        self.time_slider.setTickPosition(QSlider.TicksBelow)  # 슬라이더 아래에 눈금을 표시합니다.
        self.time_slider.setTickInterval(1)  # 눈금 간격을 1로 설정합니다.
        wait_layout.addWidget(self.time_slider, 1)  # 슬라이더를 레이아웃에 추가하고, 늘어나는 공간을 차지하도록 설정합니다.
        self.slider_value_label = QLabel("0.5초", self)  # 슬라이더 값을 표시할 레이블을 생성합니다.
        wait_layout.addWidget(self.slider_value_label)  # 레이블을 레이아웃에 추가합니다.
        settings_layout.addLayout(wait_layout)  # 대기 시간 설정 레이아웃을 설정 그룹 레이아웃에 추가합니다.

        # 슬라이더 값 변경 시 이벤트를 연결합니다.
        self.time_slider.valueChanged.connect(self.update_slider_label)  # 슬라이더 값이 변경될 때 update_slider_label 메서드를 호출합니다.

        settings_group.setLayout(settings_layout)  # 설정 그룹에 레이아웃을 설정합니다.
        self.capture_layout.addWidget(settings_group)  # 설정 그룹을 캡처 탭 레이아웃에 추가합니다.

        self.capture_layout.addSpacing(10)  # 그룹 사이에 10픽셀의 간격을 추가합니다.

        # 캡처할 창 선택 그룹을 생성합니다.
        window_group = QGroupBox("캡처할 창 선택")  # "캡처할 창 선택"이라는 제목의 그룹 박스를 생성합니다.
        window_layout = QVBoxLayout()  # 창 선택 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.

        self.refresh_windows_button = QPushButton("창 목록 새로고침", self)  # "창 목록 새로고침" 버튼을 생성합니다.
        self.refresh_windows_button.clicked.connect(self.refresh_window_list)  # 버튼 클릭 시 refresh_window_list 메서드를 호출합니다.
        window_layout.addWidget(self.refresh_windows_button)  # 버튼을 레이아웃에 추가합니다.

        self.window_list = QListWidget(self)  # 창 목록을 표시할 리스트 위젯을 생성합니다.
        self.window_list.setFixedHeight(220)  # 리스트 위젯의 높이를 220픽셀로 고정합니다.
        self.window_list.itemClicked.connect(self.select_window_from_list)  # 항목 클릭 시 select_window_from_list 메서드를 호출합니다.
        window_layout.addWidget(self.window_list)  # 리스트 위젯을 레이아웃에 추가합니다.

        window_group.setLayout(window_layout)  # 창 선택 그룹에 레이아웃을 설정합니다.
        self.capture_layout.addWidget(window_group)  # 창 선택 그룹을 캡처 탭 레이아웃에 추가합니다.

        self.capture_layout.addSpacing(10)  # 그룹 사이에 10픽셀의 간격을 추가합니다.

        # 마우스 클릭 위치 그룹을 생성합니다.
        mouse_group = QGroupBox("마우스 클릭 위치")  # "마우스 클릭 위치"라는 제목의 그룹 박스를 생성합니다.
        mouse_layout = QVBoxLayout()  # 마우스 클릭 위치 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.

        self.set_mouse_position_button = QPushButton("마우스 클릭 위치 설정", self)  # "마우스 클릭 위치 설정" 버튼을 생성합니다.
        self.set_mouse_position_button.clicked.connect(self.set_mouse_position)  # 버튼 클릭 시 set_mouse_position 메서드를 호출합니다.
        mouse_layout.addWidget(self.set_mouse_position_button)  # 버튼을 레이아웃에 추가합니다.

        self.position_label = QLabel("클릭할 위치: (0, 0)", self)  # 클릭 위치를 표시할 레이블을 생성합니다.
        mouse_layout.addWidget(self.position_label)  # 레이블을 레이아웃에 추가합니다.

        mouse_group.setLayout(mouse_layout)  # 마우스 클릭 위치 그룹에 레이아웃을 설정합니다.
        self.capture_layout.addWidget(mouse_group)  # 마우스 클릭 위치 그룹을 캡처 탭 레이아웃에 추가합니다.

        self.capture_layout.addSpacing(10)  # 그룹 사이에 10픽셀의 간격을 추가합니다.

        # 매크로 진행 상황을 표시할 UI를 생성합니다.
        self.macro_progress_layout = QHBoxLayout()  # 매크로 진행 상황을 위한 수평 박스 레이아웃을 생성합니다.
        self.macro_progress_bar = QProgressBar(self)  # 진행 상황을 표시할 프로그레스 바를 생성합니다.
        self.macro_progress_bar.setVisible(False)  # 초기에는 프로그레스 바를 숨깁니다.
        self.macro_progress_label = QLabel("0/0", self)  # 진행 상황을 텍스트로 표시할 레이블을 생성합니다.
        self.macro_progress_label.setVisible(False)  # 초기에는 레이블을 숨깁니다.
        self.macro_progress_label.setStyleSheet("font-weight: bold;")  # 레이블의 텍스트를 굵게 설정합니다.
        self.macro_progress_layout.addWidget(self.macro_progress_bar)  # 프로그레스 바를 레이아웃에 추가합니다.
        self.macro_progress_layout.addWidget(self.macro_progress_label)  # 레이블을 레이아웃에 추가합니다.
        self.capture_layout.addLayout(self.macro_progress_layout)  # 매크로 진행 상황 레이아웃을 캡처 탭 레이아웃에 추가합니다.

        self.capture_layout.addSpacing(10)  # 매크로 진행 상황과 시작 버튼 사이에 10픽셀의 간격을 추가합니다.

        # 매크로 시작 버튼을 생성합니다.
        self.start_button = QPushButton("매크로 시작", self)  # "매크로 시작" 버튼을 생성합니다.
        self.start_button.clicked.connect(self.start_macro)  # 버튼 클릭 시 start_macro 메서드를 호출합니다.
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #007bff;  /* 버튼 배경색을 파란색으로 설정합니다. */
                color: white;  /* # 버튼 텍스트 색상을 흰색으로 설정합니다. */
                padding: 10px;  /* # 버튼 내부에 10픽셀의 패딩을 추가합니다. */
                border-radius: 4px;  /* # 버튼 모서리를 4픽셀 둥글게 만듭니다. */
            }
            QPushButton:hover {
                background-color: #0056b3;  /* # 마우스를 올렸을 때 버튼 배경색을 약간 어둡게 변경합니다. */
            }
        """)  # 버튼의 스타일을 설정합니다.
        self.capture_layout.addWidget(self.start_button)  # 매크로 시작 버튼을 캡처 탭 레이아웃에 추가합니다.

        self.capture_layout.addStretch(1)  # 남은 공간을 채웁니다.

        self.click_position = (0, 0)  # 클릭 위치를 (0, 0)으로 초기화합니다.

        # 캡처 탭에 레이아웃을 설정합니다.
        capture_widget = QWidget()  # 새로운 위젯을 생성합니다.
        capture_widget.setLayout(self.capture_layout)  # 위젯에 캡처 탭 레이아웃을 설정합니다.
        self.capture_tab.setLayout(QVBoxLayout())  # 캡처 탭에 수직 박스 레이아웃을 설정합니다.
        self.capture_tab.layout().addWidget(capture_widget)  # 캡처 탭 레이아웃에 위젯을 추가합니다.

        # 캡처 탭에 레이아웃을 설정합니다.
        if self.capture_tab.layout() is None:
            self.capture_tab.setLayout(QVBoxLayout())
        self.capture_tab.layout().addWidget(capture_widget)

    def update_slider_label(self, value):
        self.slider_value_label.setText(f"{value * 0.5}초")  # 슬라이더 값에 0.5를 곱하여 초 단위로 표시합니다.

    def refresh_window_list(self):
        try:
            all_windows = [win.title for win in gw.getAllWindows() if win.title]  # 모든 창의 제목을 가져옵니다.
            limited_windows = all_windows[:20]  # 최대 20개의 창만 선택합니다.

            self.window_list.clear()  # 기존 창 목록을 지웁니다.
            for title in limited_windows:
                self.window_list.addItem(title)  # 각 창 제목을 리스트 위젯에 추가합니다.
            
            if not limited_windows:
                QMessageBox.information(self, "정보", "현재 열린 창이 없습니다.")  # 열린 창이 없을 경우 메시지를 표시합니다.
        except Exception as e:
            error_message = f"창 목록을 새로고침하는 중 오류가 발생했습니다: {str(e)}"
            QMessageBox.critical(self, "오류", error_message)  # 오류 발생 시 메시지 박스를 표시합니다.
            logging.error(error_message)  # 오류 내용을 로그에 기록합니다.

    def select_window_from_list(self, item):
        self.selected_window = item.text()  # 선택한 창의 제목을 저장합니다.
        QMessageBox.information(self, "창 선택", f"선택된 창: {self.selected_window}")  # 선택한 창을 메시지 박스로 표시합니다.

    def set_mouse_position(self):
        QMessageBox.information(self, "마우스 위치 설정", "설정할 마우스 위치에 마우스를 놓고 3초 후 자동으로 위치가 설정됩니다.")  # 안내 메시지를 표시합니다.
        time.sleep(3)  # 3초 대기합니다.
        self.click_position = pyautogui.position()  # 현재 마우스 위치를 저장합니다.
        self.position_label.setText(f"클릭할 위치: {self.click_position}")  # 저장된 위치를 레이블에 표시합니다.

    def start_macro(self):
        if not self.selected_window:
            QMessageBox.warning(self, "경고", "캡처할 창을 선택해주세요.")  # 창이 선택되지 않았을 경우 경고 메시지를 표시합니다.
            return

        repeat_count = self.repeat_spinbox.value()  # 설정된 반복 횟수를 가져옵니다.
        wait_time = self.time_slider.value() * 0.5  # 설정된 대기 시간을 가져옵니다.
        position = self.click_position  # 설정된 클릭 위치를 가져옵니다.

        QMessageBox.information(self, "매크로 시작", f"2초 후 매크로가 {repeat_count}회 실행됩니다. \n선택된 창: {self.selected_window}")  # 매크로 시작 안내 메시지를 표시합니다.
        time.sleep(2)  # 2초 대기합니다.

        try:
            window = gw.getWindowsWithTitle(self.selected_window)[0]
            window.activate()
            time.sleep(1)  # 창이 활성화되고 포커스를 받을 때까지 1초 대기

            self.macro_progress_bar.setRange(0, repeat_count)  # 프로그레스 바의 범위를 설정합니다.
            self.macro_progress_bar.setValue(0)  # 프로그레스 바 초기값을 설정합니다.
            self.macro_progress_bar.setVisible(True)  # 프로그레스 바를 표시합니다.
            self.macro_progress_label.setVisible(True)  # 프로그레스 레이블을 표시합니다.

            for i in range(repeat_count):
                time.sleep(wait_time)  # 설정된 대기 시간만큼 대기합니다.
                pyautogui.hotkey('alt', 'f1')  # Alt+F1 키를 입력합니다.
                time.sleep(0.1)  # 0.1초 대기합니다.
                pyautogui.click(position)  # 설정된 위치를 클릭합니다.
                
                self.macro_progress_bar.setValue(i + 1)  # 프로그레스 바를 업데이트합니다.
                self.macro_progress_label.setText(f"{i+1}/{repeat_count}")  # 프로그레스 레이블을 업데이트합니다.
                QApplication.processEvents()  # GUI 이벤트를 처리합니다.

            self.macro_progress_bar.setVisible(False)  # 프로그레스 바를 숨깁니다.
            self.macro_progress_label.setVisible(False)  # 프로그레스 레이블을 숨깁니다.
            QMessageBox.information(self, "매크로 완료", "매크로 실행이 완료되었습니다.")  # 매크로 완료 메시지를 표시합니다.
        except Exception as e:
                self.macro_progress_bar.setVisible(False)  # 오류 발생 시 프로그레스 바를 숨깁니다.
                self.macro_progress_label.setVisible(False)  # 오류 발생 시 프로그레스 레이블을 숨깁니다.
                error_msg = f"매크로 실행 중 오류가 발생했습니다: {str(e)}\n\n"  # 오류 메시지를 생성합니다.
                error_msg += traceback.format_exc()  # 오류의 상세 traceback을 메시지에 추가합니다.
                QMessageBox.critical(self, "오류", error_msg)  # 오류 메시지를 대화상자로 표시합니다.
                logging.error(error_msg)  # 오류 내용을 로그에 기록합니다.

    def setup_pdf_tab(self):
        pdf_layout = QHBoxLayout()  # PDF 탭의 메인 레이아웃을 수평 박스 레이아웃으로 설정합니다.
        pdf_layout.setContentsMargins(0, 0, 0, 0)  # 레이아웃의 여백을 모두 0으로 설정합니다.
        pdf_layout.setSpacing(0)  # 레이아웃 내 위젯 간 간격을 0으로 설정합니다.

        # 왼쪽 레이아웃
        left_widget = QWidget()  # 왼쪽 위젯을 생성합니다.
        left_main_layout = QVBoxLayout(left_widget)  # 왼쪽 위젯의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        left_main_layout.setContentsMargins(10, 10, 10, 10)  # 레이아웃의 여백을 모든 방향으로 10픽셀로 설정합니다.

        # 상단 위젯들을 위한 서브 레이아웃
        top_widget = QWidget()  # 상단 위젯을 생성합니다.
        left_layout = QVBoxLayout(top_widget)  # 상단 위젯의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        left_layout.setContentsMargins(0, 0, 0, 0)  # 레이아웃의 여백을 모두 0으로 설정합니다.

        # 버튼 스타일 정의
        button_style = """
            QPushButton {
                background-color: #4CAF50;  /* # 버튼 배경색을 초록색으로 설정합니다. */
                color: white;  /* # 버튼 텍스트 색상을 흰색으로 설정합니다. */
                padding: 5px;  /* # 버튼 내부에 5픽셀의 패딩을 추가합니다. */
                border: none;  /* # 버튼 테두리를 없앱니다. */
                border-radius: 3px;  /* # 버튼 모서리를 3픽셀 둥글게 만듭니다. */
                height: 20px;  /* # 버튼 높이를 20픽셀로 설정합니다. */
                font-size: 9pt;  /* # 버튼 텍스트 크기를 9포인트로 설정합니다. */
            }
            QPushButton:hover {
                background-color: #45a049;  /* # 마우스를 올렸을 때 버튼 배경색을 약간 어둡게 변경합니다. */
            }
        """

        # 폴더 선택 그룹
        folder_group = QGroupBox("폴더 선택")  # "폴더 선택"이라는 제목의 그룹 박스를 생성합니다.
        folder_layout = QVBoxLayout()  # 폴더 선택 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        
        self.folder_label = QLabel(f"{self.default_path}", self)  # 현재 선택된 폴더 경로를 표시할 레이블을 생성합니다.
        self.folder_label.setStyleSheet("background-color: white; padding: 5px; border: 1px solid #ccc;")  # 레이블의 스타일을 설정합니다.
        folder_layout.addWidget(self.folder_label)  # 레이블을 레이아웃에 추가합니다.
        
        self.select_folder_button = QPushButton("폴더 선택", self)  # "폴더 선택" 버튼을 생성합니다.
        self.select_folder_button.clicked.connect(self.select_folder)  # 버튼 클릭 시 select_folder 메서드를 호출합니다.
        self.select_folder_button.setStyleSheet(button_style)  # 버튼의 스타일을 설정합니다.
        folder_layout.addWidget(self.select_folder_button)  # 버튼을 레이아웃에 추가합니다.

        self.change_default_button = QPushButton("기본 경로 변경", self)  # "기본 경로 변경" 버튼을 생성합니다.
        self.change_default_button.clicked.connect(self.change_default_path)  # 버튼 클릭 시 change_default_path 메서드를 호출합니다.
        self.change_default_button.setStyleSheet(button_style)  # 버튼의 스타일을 설정합니다.
        folder_layout.addWidget(self.change_default_button)  # 버튼을 레이아웃에 추가합니다.

        folder_group.setLayout(folder_layout)  # 폴더 선택 그룹에 레이아웃을 설정합니다.
        left_layout.addWidget(folder_group)  # 폴더 선택 그룹을 왼쪽 레이아웃에 추가합니다.

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.

        # 크롭 설정 그룹
        crop_group = QGroupBox("크롭 설정")  # "크롭 설정"이라는 제목의 그룹 박스를 생성합니다.
        crop_layout = QVBoxLayout()  # 크롭 설정 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        
        self.coord_label = QLabel("크롭 좌표: ", self)  # 크롭 좌표를 표시할 레이블을 생성합니다.
        crop_layout.addWidget(self.coord_label)  # 레이블을 레이아웃에 추가합니다.

        for position, name in [("위쪽", "top"), ("아래쪽", "bottom"), ("왼쪽", "left"), ("오른쪽", "right")]:
            layout = QHBoxLayout()  # 각 위치별로 수평 박스 레이아웃을 생성합니다.
            layout.addWidget(QLabel(f"{position.rjust(4)} : ", self))  # 위치 레이블을 생성하고 레이아웃에 추가합니다.
            spin = QSpinBox(self)  # 스핀 박스를 생성합니다.
            spin.setRange(0, 10000)  # 스핀 박스의 범위를 0부터 10000까지로 설정합니다.
            spin.valueChanged.connect(self.update_crop_from_spinbox)  # 값 변경 시 update_crop_from_spinbox 메서드를 호출합니다.
            setattr(self, f"{name}_spin", spin)  # 스핀 박스 객체를 클래스의 속성으로 설정합니다.
            layout.addWidget(spin)  # 스핀 박스를 레이아웃에 추가합니다.
            crop_layout.addLayout(layout)  # 생성한 레이아웃을 크롭 설정 레이아웃에 추가합니다.

        crop_group.setLayout(crop_layout)  # 크롭 설정 그룹에 레이아웃을 설정합니다.
        left_layout.addWidget(crop_group)  # 크롭 설정 그룹을 왼쪽 레이아웃에 추가합니다.

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.

        # PDF 방향 선택을 위한 그룹 박스 추가
        orientation_group_box = QGroupBox("용지 방향 설정")
        orientation_layout = QHBoxLayout()  # 가로로 배치하기 위한 QHBoxLayout 사용

        # PDF 방향 선택을 위한 라디오 버튼 추가
        self.orientation_group = QButtonGroup(self)
        self.portrait_radio = QRadioButton("세로")
        self.landscape_radio = QRadioButton("가로")
        self.orientation_group.addButton(self.portrait_radio)
        self.orientation_group.addButton(self.landscape_radio)
        self.portrait_radio.setChecked(True)  # 기본값을 세로로 설정

        # 그룹 박스 레이아웃에 라디오 버튼 추가 및 간격 추가
        orientation_layout.addWidget(self.portrait_radio)

        # 간격을 위한 SpacerItem 추가 (숫자로 간격 지정)
        spacer = QSpacerItem(50, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)  # 40px 가로 간격
        orientation_layout.addItem(spacer)

        orientation_layout.addWidget(self.landscape_radio)

        # 라디오 버튼을 그룹 내에서 좌우 중앙에 정렬
        orientation_layout.setAlignment(Qt.AlignCenter)

        # 그룹 박스에 레이아웃 설정
        orientation_group_box.setLayout(orientation_layout)

        # 전체 레이아웃에 그룹 박스 추가
        left_layout.addWidget(orientation_group_box)

        # 레이아웃을 위젯에 설정
        self.setLayout(left_layout)

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.
        
        # 압축 설정 추가
        compression_layout = QHBoxLayout()  # 압축 설정을 위한 수평 박스 레이아웃을 생성합니다.
        compression_layout.addWidget(QLabel("압축 수준:"))  # "압축 수준:" 레이블을 생성하고 레이아웃에 추가합니다.
        self.compression_slider = QSlider(Qt.Horizontal)  # 수평 슬라이더를 생성합니다.
        self.compression_slider.setRange(0, 100)  # 슬라이더의 범위를 0부터 100까지로 설정합니다.
        self.compression_slider.setValue(85)  # 슬라이더의 초기값을 85로 설정합니다.
        self.compression_slider.setTickPosition(QSlider.TicksBelow)  # 슬라이더 아래에 눈금을 표시합니다.
        self.compression_slider.setTickInterval(10)  # 눈금 간격을 10으로 설정합니다.
        compression_layout.addWidget(self.compression_slider)  # 슬라이더를 레이아웃에 추가합니다.
        self.compression_value_label = QLabel("85")  # 압축 수준 값을 표시할 레이블을 생성합니다.
        compression_layout.addWidget(self.compression_value_label)  # 레이블을 레이아웃에 추가합니다.
        left_layout.addLayout(compression_layout)  # 압축 설정 레이아웃을 왼쪽 레이아웃에 추가합니다.

        self.compression_slider.valueChanged.connect(self.update_compression_label)  # 슬라이더 값 변경 시 update_compression_label 메서드를 호출합니다.

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.

        # PDF 생성 버튼
        self.create_pdf_button = QPushButton("PDF 생성", self)  # "PDF 생성" 버튼을 생성합니다.
        self.create_pdf_button.clicked.connect(self.create_pdf)  # 버튼 클릭 시 create_pdf 메서드를 호출합니다.
        self.create_pdf_button.setStyleSheet(button_style)  # 버튼의 스타일을 설정합니다.
        left_layout.addWidget(self.create_pdf_button)  # 버튼을 왼쪽 레이아웃에 추가합니다.

        left_layout.addSpacing(10)  # 그룹 박스 사이에 10픽셀의 간격을 추가합니다.

        # PDF 자동 열기 옵션
        self.auto_open_checkbox = QCheckBox("PDF 생성 후 자동으로 열기", self)  # "PDF 생성 후 자동으로 열기" 체크박스를 생성합니다.
        left_layout.addWidget(self.auto_open_checkbox)  # 체크박스를 왼쪽 레이아웃에 추가합니다.

        left_layout.addSpacing(100)  # 그룹 박스 사이에 100픽셀의 간격을 추가합니다.

        # 상단 위젯들을 left_main_layout에 추가
        left_main_layout.addWidget(top_widget)  # 상단 위젯을 왼쪽 메인 레이아웃에 추가합니다.
        left_main_layout.addStretch(1)  # 상단 위젯과 프로그레스 바 사이에 신축성 있는 공간을 추가합니다.

        # 프로그레스 바와 라벨 (화면 아래쪽에 고정)
        progress_group = QGroupBox("진행 상황")  # "진행 상황"이라는 제목의 그룹 박스를 생성합니다.
        progress_layout = QVBoxLayout()  # 진행 상황 그룹의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        
        self.progress_bar = QProgressBar(self)  # 프로그레스 바를 생성합니다.
        self.progress_bar.setVisible(False)  # 초기에는 프로그레스 바를 숨깁니다.
        progress_layout.addWidget(self.progress_bar)  # 프로그레스 바를 레이아웃에 추가합니다.
        
        self.progress_label = QLabel("0/0", self)  # 진행 상황을 표시할 레이블을 생성합니다.
        self.progress_label.setVisible(False)  # 초기에는 레이블을 숨깁니다.
        self.progress_label.setAlignment(Qt.AlignCenter)  # 레이블의 텍스트를 중앙 정렬합니다.
        progress_layout.addWidget(self.progress_label)  # 레이블을 레이아웃에 추가합니다.
        
        progress_group.setLayout(progress_layout)  # 진행 상황 그룹에 레이아웃을 설정합니다.
        left_main_layout.addWidget(progress_group)  # 진행 상황 그룹을 왼쪽 메인 레이아웃에 추가합니다.

        # 왼쪽 메뉴의 고정 너비 설정
        left_widget.setFixedWidth(380)  # 왼쪽 위젯의 너비를 380픽셀로 고정합니다.

        # 오른쪽 레이아웃 (이미지 미리보기)
        right_widget = QWidget()  # 오른쪽 위젯을 생성합니다.
        preview_layout = QVBoxLayout(right_widget)  # 오른쪽 위젯의 레이아웃을 수직 박스 레이아웃으로 설정합니다.
        preview_layout.setContentsMargins(0, 0, 0, 0)  # 레이아웃의 여백을 모두 0으로 설정합니다.

        self.scroll_area = QScrollArea(self)  # 스크롤 영역을 생성합니다.
        self.scroll_area.setWidgetResizable(True)  # 스크롤 영역 내 위젯의 크기를 조절 가능하게 설정합니다.
        self.image_widget = ImageWidget(self)  # ImageWidget 인스턴스를 생성합니다.
        self.scroll_area.setWidget(self.image_widget)  # 스크롤 영역에 ImageWidget을 설정합니다.
        preview_layout.addWidget(self.scroll_area)  # 스크롤 영역을 미리보기 레이아웃에 추가합니다.
        
        button_layout = QHBoxLayout()  # 버튼을 위한 수평 박스 레이아웃을 생성합니다.
        self.prev_button = QPushButton("< 이전 이미지", self)  # "< 이전 이미지" 버튼을 생성합니다.
        self.next_button = QPushButton("다음 이미지 >", self)  # "다음 이미지 >" 버튼을 생성합니다.
        self.prev_button.clicked.connect(self.show_previous_image)  # 이전 이미지 버튼 클릭 시 show_previous_image 메서드를 호출합니다.
        self.next_button.clicked.connect(self.show_next_image)  # 다음 이미지 버튼 클릭 시 show_next_image 메서드를 호출합니다.
        self.fit_to_view_button = QPushButton("창에 맞추기", self)  # "창에 맞추기" 버튼을 생성합니다.
        self.fit_to_view_button.clicked.connect(self.fit_image_to_view)  # 창에 맞추기 버튼 클릭 시 fit_image_to_view 메서드를 호출합니다.
        button_layout.addWidget(self.prev_button)  # 이전 이미지 버튼을 버튼 레이아웃에 추가합니다.
        button_layout.addWidget(self.next_button)  # 다음 이미지 버튼을 버튼 레이아웃에 추가합니다.
        button_layout.addWidget(self.fit_to_view_button)  # 창에 맞추기 버튼을 버튼 레이아웃에 추가합니다.
        preview_layout.addLayout(button_layout)  # 버튼 레이아웃을 미리보기 레이아웃에 추가합니다.

        # 메인 레이아웃에 왼쪽과 오른쪽 위젯 추가
        pdf_layout.addWidget(left_widget)  # 왼쪽 위젯을 PDF 레이아웃에 추가합니다.
        pdf_layout.addWidget(right_widget, 1)  # 오른쪽 위젯을 PDF 레이아웃에 추가하고, 늘어나는 공간을 차지하도록 설정합니다.

        pdf_widget = QWidget()  # PDF 위젯을 생성합니다.
        pdf_widget.setLayout(pdf_layout)  # PDF 위젯에 PDF 레이아웃을 설정합니다.
        self.pdf_tab.setLayout(QVBoxLayout())  # PDF 탭에 수직 박스 레이아웃을 설정합니다.
        self.pdf_tab.layout().addWidget(pdf_widget)  # PDF 탭 레이아웃에 PDF 위젯을 추가합니다.

        if self.pdf_tab.layout() is None:
            self.pdf_tab.setLayout(QVBoxLayout())
        self.pdf_tab.layout().addWidget(pdf_widget)

    def update_compression_label(self, value):
        self.compression_value_label.setText(str(value))  # 압축 수준 슬라이더의 값을 레이블에 표시합니다.

    def on_tab_changed(self, index):
        if self.tab_widget.tabText(index) == "PDF 생성":  # 현재 선택된 탭이 "PDF 생성" 탭인 경우
            if not self.is_pdf_tab_initialized:  # PDF 탭이 초기화되지 않았다면
                self.initialize_pdf_tab()  # PDF 탭을 초기화합니다.
            self.adjust_image_size()  # 이미지 크기를 조정합니다.
            self.pdf_tab.layout().update()  # PDF 탭의 레이아웃을 업데이트합니다.

    def adjust_image_size(self):
        if hasattr(self, 'image_widget') and self.image_widget.pixmap:  # 이미지 위젯과 픽스맵이 존재하는 경우
            self.image_widget.fit_image()  # 이미지를 위젯에 맞게 조정합니다.
            self.image_widget.update()  # 이미지 위젯을 업데이트합니다.

    def adjust_scroll_bar(self):
        if self.image_widget.pixmap:  # 이미지 픽스맵이 존재하는 경우
            self.scroll_area.setWidgetResizable(True)  # 스크롤 영역 내 위젯의 크기를 조절 가능하게 설정합니다.
            self.scroll_area.updateGeometry()  # 스크롤 영역의 지오메트리를 업데이트합니다.

    def update_crop_coordinates(self, rect):
        pixmap_rect = QRect(self.image_widget.mapToPixmap(rect.topLeft()),
                            self.image_widget.mapToPixmap(rect.bottomRight()))  # 선택 영역을 픽스맵 좌표로 변환합니다.
        self.crop_rect = pixmap_rect  # 크롭 영역을 업데이트합니다.
        self.left_spin.setValue(pixmap_rect.left())  # 왼쪽 좌표를 스핀 박스에 설정합니다.
        self.top_spin.setValue(pixmap_rect.top())  # 위쪽 좌표를 스핀 박스에 설정합니다.
        self.right_spin.setValue(pixmap_rect.right())  # 오른쪽 좌표를 스핀 박스에 설정합니다.
        self.bottom_spin.setValue(pixmap_rect.bottom())  # 아래쪽 좌표를 스핀 박스에 설정합니다.
        self.coord_label.setText(f"크롭 좌표: ({pixmap_rect.left()}, {pixmap_rect.top()}, {pixmap_rect.right()}, {pixmap_rect.bottom()})")  # 크롭 좌표를 레이블에 표시합니다.

    def update_crop_from_spinbox(self):
        left = self.left_spin.value()  # 왼쪽 스핀 박스의 값을 가져옵니다.
        top = self.top_spin.value()  # 위쪽 스핀 박스의 값을 가져옵니다.
        right = self.right_spin.value()  # 오른쪽 스핀 박스의 값을 가져옵니다.
        bottom = self.bottom_spin.value()  # 아래쪽 스핀 박스의 값을 가져옵니다.
        self.crop_rect = QRect(left, top, right - left, bottom - top)  # 크롭 영역을 업데이트합니다.
        self.image_widget.rubberband = QRect(
            self.image_widget.mapFromPixmap(QPoint(left, top)),
            self.image_widget.mapFromPixmap(QPoint(right, bottom))
        ).normalized()  # 선택 영역을 위젯 좌표로 변환하여 업데이트합니다.
        self.image_widget.update()  # 이미지 위젯을 업데이트합니다.
        self.coord_label.setText(f"크롭 좌표: ({left}, {top}, {right}, {bottom})")  # 크롭 좌표를 레이블에 표시합니다.

    def create_pdf(self):
        try:
            if not hasattr(self, 'crop_rect') or self.crop_rect.isNull():
                QMessageBox.warning(self, "경고", "크롭 영역을 선택해주세요.")
                return
        
            if not hasattr(self, 'cropper_folder') or not self.cropper_folder:
                QMessageBox.warning(self, "경고", "먼저 폴더를 선택해주세요.")
                return
        
            current_time = datetime.now().strftime("%y%m_%H%M_%S")
            pdf_name = f"cropped_ebook_{current_time}.pdf"
            pdf_path = os.path.join(self.cropper_folder, pdf_name)
        
            images = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            if not images:
                QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")
                return

            total_steps = len(images)
        
            # 방향 설정
            orientation = self.orientation_group.checkedButton().text()
            if orientation == "가로":
                pagesize = landscape(A4)
            else:
                pagesize = A4

            c = canvas.Canvas(pdf_path, pagesize=pagesize)
            page_width, page_height = pagesize

            self.progress_bar.setRange(0, total_steps)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_label.setVisible(True)

            batch_size = 10
        
            with tempfile.TemporaryDirectory() as temp_dir:
                for i in range(0, len(images), batch_size):
                    batch = images[i:i+batch_size]
                
                    for img_file in batch:
                        try:
                            img_path = os.path.join(self.image_folder, img_file)
                            img = Image.open(img_path)
                            cropped_img = img.crop((self.crop_rect.left(), self.crop_rect.top(),
                                                    self.crop_rect.right(), self.crop_rect.bottom()))
                        
                            temp_img_path = os.path.join(temp_dir, f"temp_{img_file}")
                            compression_level = self.compression_slider.value()
                            cropped_img.save(temp_img_path, format='JPEG', quality=compression_level, optimize=True)
                        
                            img_reader = ImageReader(temp_img_path)
                            img_width, img_height = cropped_img.size
                            width_ratio = page_width / img_width
                            height_ratio = page_height / img_height
                            scale_factor = min(width_ratio, height_ratio)
                        
                            new_width = img_width * scale_factor
                            new_height = img_height * scale_factor
                        
                            x_centered = (page_width - new_width) / 2
                            y_centered = (page_height - new_height) / 2
                        
                            c.drawImage(img_reader, x_centered, y_centered, width=new_width, height=new_height)
                            c.showPage()
                        
                            self.progress_bar.setValue(i + batch.index(img_file) + 1)
                            self.progress_label.setText(f"처리 중: {i + batch.index(img_file) + 1}/{total_steps}\n{img_file}")
                            QApplication.processEvents()

                        except Exception as e:
                            logging.error(f"이미지 처리 중 오류 발생 {img_file}: {str(e)}")
                            QMessageBox.warning(self, "경고", f"이미지 처리 중 오류 발생: {img_file}\n{str(e)}")

                c.save()
            
            compressed_pdf_path = self.compress_pdf(pdf_path)

            self.progress_bar.setVisible(False)
            self.progress_label.setVisible(False)
            QMessageBox.information(self, "완료", f"PDF 생성 및 압축이 완료되었습니다.\n저장 위치: {compressed_pdf_path}")

            if self.auto_open_checkbox.isChecked():
                self.open_pdf(compressed_pdf_path)
            
        except Exception as e:
            error_msg = f"PDF 생성 중 오류가 발생했습니다: {str(e)}\n\n"
            error_msg += traceback.format_exc()
            if hasattr(self, 'progress_bar'):
                self.progress_bar.setVisible(False)
            if hasattr(self, 'progress_label'):
                self.progress_label.setVisible(False)
            QMessageBox.critical(self, "오류", error_msg)
            logging.error(error_msg)

    def compress_pdf(self, input_path, output_path=None):
        if output_path is None:
            output_path = input_path  # 출력 경로가 지정되지 않은 경우 입력 경로를 사용합니다.
        
        reader = PdfReader(input_path)  # PDF 파일을 읽습니다.
        writer = PdfWriter()  # 새로운 PDF 작성자를 생성합니다.

        for page in reader.pages:
            page.compress_content_streams()  # 각 페이지의 내용을 압축합니다. (CPU 집약적인 작업입니다)
            writer.add_page(page)  # 압축된 페이지를 새 PDF에 추가합니다.

        with open(output_path, "wb") as f:
            writer.write(f)  # 압축된 PDF를 파일로 저장합니다.

        return output_path  # 압축된 PDF의 경로를 반환합니다.

    def open_pdf(self, pdf_path):
        try:
            if sys.platform.startswith('darwin'):  # macOS인 경우
                subprocess.run(['open', pdf_path])  # 'open' 명령어로 PDF를 엽니다.
            elif sys.platform.startswith('win'):  # Windows인 경우
                os.startfile(pdf_path)  # os.startfile()로 PDF를 엽니다.
            else:  # Linux 등 기타 OS인 경우
                subprocess.run(['xdg-open', pdf_path])  # 'xdg-open' 명령어로 PDF를 엽니다.
        except Exception as e:
            QMessageBox.warning(self, "경고", f"PDF 파일을 열 수 없습니다: {str(e)}")  # 오류 메시지를 표시합니다.

    def initialize_pdf_tab(self):
        self.initialize_folders()  # 폴더를 초기화합니다.
        self.move_files_to_image_folder()  # 이미지 파일을 이동합니다.
        self.load_first_image()  # 첫 번째 이미지를 로드합니다.
        self.is_pdf_tab_initialized = True  # PDF 탭 초기화 완료 플래그를 설정합니다.

    def initialize_folders(self):
        try:
            if not os.path.exists(self.image_folder):
                os.makedirs(self.image_folder)  # 이미지 폴더가 없으면 생성합니다.
            if not os.path.exists(self.cropper_folder):
                os.makedirs(self.cropper_folder)  # 크로퍼 폴더가 없으면 생성합니다.
            self.move_files_to_image_folder()  # 이미지 파일을 이동합니다.
        except Exception as e:
            error_msg = f"폴더 초기화 중 오류가 발생했습니다: {str(e)}\n\n"  # 오류 메시지를 생성합니다.
            error_msg += traceback.format_exc()  # 상세한 오류 정보를 추가합니다.
            QMessageBox.critical(self, "오류", error_msg)  # 오류 메시지를 표시합니다.
            logging.error(error_msg)  # 오류를 로그에 기록합니다.

    def move_files_to_image_folder(self):
        image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')  # 이미지 파일 확장자 목록입니다.
        moved_files = 0  # 이동된 파일 수를 추적합니다.
        for filename in os.listdir(self.base_folder):  # 기본 폴더의 모든 파일에 대해
            if filename.lower().endswith(image_extensions):  # 이미지 파일인 경우
                src_path = os.path.join(self.base_folder, filename)  # 원본 파일 경로
                dst_path = os.path.join(self.image_folder, filename)  # 대상 파일 경로
                try:
                    shutil.move(src_path, dst_path)  # 파일을 이동합니다.
                    moved_files += 1  # 이동된 파일 수를 증가시킵니다.
                    logging.info(f"Moved file: {filename}")  # 로그에 기록합니다.
                except Exception as e:
                    logging.error(f"Error moving file {filename}: {str(e)}")  # 오류를 로그에 기록합니다.
        
        if moved_files > 0:
            logging.info(f"{moved_files}개의 이미지 파일이 Image 폴더로 이동되었습니다.")  # 이동된 파일 수를 로그에 기록합니다.

    def load_first_image(self):
        try:
            self.image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]  # 이미지 파일 목록을 가져옵니다.
            if not self.image_files:  # 이미지 파일이 없는 경우
                QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")  # 경고 메시지를 표시합니다.
                return

            self.current_image_index = 0  # 현재 이미지 인덱스를 0으로 설정합니다.
            self.load_image(self.current_image_index)  # 첫 번째 이미지를 로드합니다.
        except Exception as e:
            error_msg = f"이미지 로드 중 오류가 발생했습니다: {str(e)}\n\n"  # 오류 메시지를 생성합니다.
            error_msg += traceback.format_exc()  # 상세한 오류 정보를 추가합니다.
            QMessageBox.critical(self, "오류", error_msg)  # 오류 메시지를 표시합니다.
            logging.error(error_msg)  # 오류를 로그에 기록합니다.

    def load_image(self, index):
        if 0 <= index < len(self.image_files):  # 유효한 인덱스인지 확인합니다.
            try:
                image_path = os.path.join(self.image_folder, self.image_files[index])  # 이미지 파일 경로를 생성합니다.
                self.current_image = QPixmap(image_path)  # 이미지를 QPixmap으로 로드합니다.
                if self.current_image.isNull():  # 이미지 로드에 실패한 경우
                    raise Exception(f"이미지를 불러올 수 없습니다: {image_path}")  # 예외를 발생시킵니다.
                self.image_widget.setPixmap(self.current_image)  # 이미지 위젯에 이미지를 설정합니다.
                self.fit_image_to_view()  # 이미지를 뷰에 맞게 조정합니다.

                for spin in [self.left_spin, self.top_spin, self.right_spin, self.bottom_spin]:
                    spin.setMaximum(max(self.current_image.width(), self.current_image.height()))  # 스핀 박스의 최대값을 이미지 크기에 맞게 설정합니다.

                self.prev_button.setEnabled(index > 0)  # 이전 버튼 활성화 여부를 설정합니다.
                self.next_button.setEnabled(index < len(self.image_files) - 1)  # 다음 버튼 활성화 여부를 설정합니다.
            except Exception as e:
                error_msg = f"이미지 로드 중 오류가 발생했습니다: {str(e)}\n\n"  # 오류 메시지를 생성합니다.
                error_msg += traceback.format_exc()  # 상세한 오류 정보를 추가합니다.
                QMessageBox.critical(self, "오류", error_msg)  # 오류 메시지를 표시합니다.
                logging.error(error_msg)  # 오류를 로그에 기록합니다.

    def fit_image_to_view(self):
        if self.image_widget.pixmap:  # 이미지가 로드되어 있는 경우
            self.image_widget.fit_image_to_view()  # 이미지 위젯의 fit_image_to_view 메서드를 호출합니다.
            self.adjust_scroll_bar()  # 스크롤바를 조정합니다.

    def show_previous_image(self):
        if hasattr(self, 'current_image_index'):  # current_image_index 속성이 있는 경우
            self.current_image_index = max(0, self.current_image_index - 1)  # 이전 이미지 인덱스를 계산합니다.
            self.load_image(self.current_image_index)  # 이전 이미지를 로드합니다.

    def show_next_image(self):
        if hasattr(self, 'current_image_index'):  # current_image_index 속성이 있는 경우
            self.current_image_index = min(len(self.image_files) - 1, self.current_image_index + 1)  # 다음 이미지 인덱스를 계산합니다.
            self.load_image(self.current_image_index)  # 다음 이미지를 로드합니다.

    def select_folder(self):
        selected_folder = QFileDialog.getExistingDirectory(self, "폴더 선택", self.default_path)  # 폴더 선택 대화상자를 엽니다.
        if selected_folder:  # 폴더가 선택된 경우
            self.base_folder = selected_folder  # 기본 폴더를 업데이트합니다.
            self.folder_label.setText(f"선택된 폴더: {self.base_folder}")  # 폴더 레이블을 업데이트합니다.
            self.image_folder = os.path.join(self.base_folder, "Image")  # 이미지 폴더 경로를 업데이트합니다.
            self.cropper_folder = os.path.join(self.base_folder, "Cropper")  # 크로퍼 폴더 경로를 업데이트합니다.
            self.initialize_pdf_tab()  # PDF 탭을 초기화합니다.

    def change_default_path(self):
        new_default_path = QFileDialog.getExistingDirectory(self, "새 기본 경로 선택", self.default_path)  # 새 기본 경로 선택 대화상자를 엽니다.
        if new_default_path:  # 새 경로가 선택된 경우
            self.default_path = new_default_path  # 기본 경로를 업데이트합니다.
            QMessageBox.information(self, "기본 경로 변경", f"새 기본 경로가 설정되었습니다: {self.default_path}")  # 정보 메시지를 표시합니다.
            self.folder_label.setText(f"선택된 폴더: {self.default_path}")  # 폴더 레이블을 업데이트합니다.

if __name__ == "__main__":
    app = QApplication(sys.argv)  # QApplication 인스턴스를 생성합니다.
    window = IntegratedEBookApp()  # IntegratedEBookApp 인스턴스를 생성합니다.
    window.show()  # 윈도우를 표시합니다.
    sys.exit(app.exec_())  # 애플리케이션의 이벤트 루프를 시작합니다.
