import sys
import time
import pyautogui
import os
import shutil
import logging
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QSpinBox, QSlider, QWidget, QMessageBox,
                             QFileDialog, QScrollArea, QTabWidget, QGroupBox, QComboBox)
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QRect, QPoint, QSize
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class ImageWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.pixmap = None
        self.rubberband = QRect()
        self.origin = QPoint()
        self.zoom_factor = 1.0

    def setPixmap(self, pixmap):
        self.pixmap = pixmap
        self.update()

    def paintEvent(self, event):
        if self.pixmap:
            painter = QPainter(self)
            scaled_pixmap = self.pixmap.scaled(self.size() * self.zoom_factor, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            painter.drawPixmap(self.rect(), scaled_pixmap)
            if not self.rubberband.isNull():
                painter.setPen(QPen(Qt.red, 2, Qt.SolidLine))
                painter.drawRect(self.rubberband)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.origin = event.pos()
            self.rubberband = QRect(self.origin, QSize())
            self.update()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton:
            self.rubberband = QRect(self.origin, event.pos()).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.rubberband = QRect(self.origin, event.pos()).normalized()
            self.parent.update_crop_coordinates(self.rubberband)
            self.update()

    def set_zoom(self, zoom_factor):
        self.zoom_factor = zoom_factor
        self.update()

class IntegratedEBookApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("통합 eBook 캡처 및 PDF 생성기")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 14px;
                color: #333;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #39843c;
            }
            QSpinBox, QSlider, QComboBox {
                background-color: #ffffff;
                border: 1px solid #bbb;
                padding: 5px;
                font-size: 14px;
            }
            QTabWidget::pane {
                border: 1px solid #d7d7d7;
                background-color: #ececec;
            }
            QTabBar::tab {
                background-color: #e1e1e1;
                padding: 10px;
                font-size: 16px;
            }
            QTabBar::tab:selected {
                background-color: #ececec;
                font-weight: bold;
            }
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                border: 2px solid #bbb;
                border-radius: 5px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
        """)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        # 캡처 탭
        self.capture_tab = QWidget()
        self.capture_layout = QVBoxLayout(self.capture_tab)
        self.tab_widget.addTab(self.capture_tab, "화면 캡처")

        # PDF 생성 탭
        self.pdf_tab = QWidget()
        self.pdf_layout = QVBoxLayout(self.pdf_tab)
        self.tab_widget.addTab(self.pdf_tab, "PDF 생성")

        self.setup_capture_tab()
        self.setup_pdf_tab()

        self.window_titles = []
        self.selected_window = None

    def setup_capture_tab(self):
        # 설정 그룹
        settings_group = QGroupBox("캡처 설정")
        settings_layout = QVBoxLayout()

        # 반복 횟수 설정
        repeat_layout = QHBoxLayout()
        self.repeat_label = QLabel("반복 횟수:", self)
        repeat_layout.addWidget(self.repeat_label)
        self.repeat_spinbox = QSpinBox(self)
        self.repeat_spinbox.setRange(1, 2000)
        self.repeat_spinbox.setValue(1)
        repeat_layout.addWidget(self.repeat_spinbox)
        settings_layout.addLayout(repeat_layout)

        # 대기 시간 설정
        wait_layout = QHBoxLayout()
        self.slider_label = QLabel("대기 시간(초):", self)
        wait_layout.addWidget(self.slider_label)
        self.time_slider = QSlider(Qt.Horizontal, self)
        self.time_slider.setRange(1, 20)
        self.time_slider.setValue(1)
        self.time_slider.setTickPosition(QSlider.TicksBelow)
        self.time_slider.setTickInterval(1)
        wait_layout.addWidget(self.time_slider)
        self.slider_value_label = QLabel("0.5초", self)
        wait_layout.addWidget(self.slider_value_label)
        settings_layout.addLayout(wait_layout)

        self.time_slider.valueChanged.connect(self.update_slider_label)

        settings_group.setLayout(settings_layout)
        self.capture_layout.addWidget(settings_group)

        # 창 선택 그룹
        window_group = QGroupBox("캡처할 창 선택")
        window_layout = QVBoxLayout()

        self.refresh_windows_button = QPushButton("창 목록 새로고침", self)
        self.refresh_windows_button.clicked.connect(self.refresh_window_list)
        window_layout.addWidget(self.refresh_windows_button)

        self.window_combo = QComboBox(self)
        self.window_combo.currentIndexChanged.connect(self.select_window)
        window_layout.addWidget(self.window_combo)

        window_group.setLayout(window_layout)
        self.capture_layout.addWidget(window_group)

        # 마우스 위치 그룹
        mouse_group = QGroupBox("마우스 클릭 위치")
        mouse_layout = QVBoxLayout()

        self.set_mouse_position_button = QPushButton("마우스 클릭 위치 설정", self)
        self.set_mouse_position_button.clicked.connect(self.set_mouse_position)
        mouse_layout.addWidget(self.set_mouse_position_button)

        self.position_label = QLabel("클릭할 위치: (0, 0)", self)
        mouse_layout.addWidget(self.position_label)

        mouse_group.setLayout(mouse_layout)
        self.capture_layout.addWidget(mouse_group)

        # 매크로 시작 버튼
        self.start_button = QPushButton("매크로 시작", self)
        self.start_button.clicked.connect(self.start_macro)
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                font-size: 18px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.capture_layout.addWidget(self.start_button)

        # 빈 공간 추가
        self.capture_layout.addStretch()

        # 초기 설정
        self.click_position = (0, 0)

    def setup_pdf_tab(self):
        # 폴더 선택 그룹
        folder_group = QGroupBox("폴더 선택")
        folder_layout = QVBoxLayout()

        select_layout = QHBoxLayout()
        self.select_folder_button = QPushButton("폴더 선택", self)
        self.select_folder_button.clicked.connect(self.select_folder)
        select_layout.addWidget(self.select_folder_button)
        self.folder_label = QLabel("선택된 폴더: 없음", self)
        select_layout.addWidget(self.folder_label)
        folder_layout.addLayout(select_layout)

        # 안내 메시지 추가
        self.guide_label = QLabel("'C:/Users/user/Videos/Desktop' 폴더를 선택하세요.", self)
        self.guide_label.setStyleSheet("color: #666; font-style: italic;")
        folder_layout.addWidget(self.guide_label)

        folder_group.setLayout(folder_layout)
        self.pdf_layout.addWidget(folder_group)

        # 이미지 미리보기 영역
        preview_group = QGroupBox("이미지 미리보기")
        preview_layout = QVBoxLayout()
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.image_widget = ImageWidget(self)
        self.scroll_area.setWidget(self.image_widget)
        preview_layout.addWidget(self.scroll_area)
        preview_group.setLayout(preview_layout)
        self.pdf_layout.addWidget(preview_group)

        # 크롭 설정 그룹
        crop_group = QGroupBox("크롭 설정")
        crop_layout = QVBoxLayout()

        # 좌표 레이아웃
        coord_layout = QHBoxLayout()
        self.coord_label = QLabel("크롭 좌표: ", self)
        coord_layout.addWidget(self.coord_label)

        self.left_spin = QSpinBox(self)
        self.top_spin = QSpinBox(self)
        self.right_spin = QSpinBox(self)
        self.bottom_spin = QSpinBox(self)
        
        for spin in [self.left_spin, self.top_spin, self.right_spin, self.bottom_spin]:
            spin.setRange(0, 10000)
            spin.valueChanged.connect(self.update_crop_from_spinbox)
            coord_layout.addWidget(spin)

        crop_layout.addLayout(coord_layout)
        
        # 줌 슬라이더
        zoom_layout = QHBoxLayout()
        zoom_label = QLabel("줌:", self)
        zoom_layout.addWidget(zoom_label)
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(10, 200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(self.update_zoom)
        zoom_layout.addWidget(self.zoom_slider)
        self.zoom_value_label = QLabel("100%", self)
        zoom_layout.addWidget(self.zoom_value_label)
        crop_layout.addLayout(zoom_layout)

        crop_group.setLayout(crop_layout)
        self.pdf_layout.addWidget(crop_group)

        # PDF 생성 버튼
        self.create_pdf_button = QPushButton("PDF 생성", self)
        self.create_pdf_button.clicked.connect(self.create_pdf)
        self.create_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                font-size: 18px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        self.pdf_layout.addWidget(self.create_pdf_button)

        # 초기화
        self.base_folder = ""
        self.image_folder = ""
        self.cropper_folder = ""
        self.current_image = None
        self.crop_rect = QRect()

    def update_slider_label(self, value):
        self.slider_value_label.setText(f"{value * 0.5}초")

    def refresh_window_list(self):
        self.window_titles = [win.title for win in pyautogui.getAllWindows() if win.title]
        self.window_combo.clear()
        self.window_combo.addItems(self.window_titles)

    def select_window(self, index):
        if 0 <= index < len(self.window_titles):
            self.selected_window = self.window_titles[index]
            QMessageBox.information(self, "창 선택", f"선택된 창: {self.selected_window}")

    def set_mouse_position(self):
        QMessageBox.information(self, "마우스 위치 설정", "설정할 마우스 위치에 마우스를 놓고 3초 후 자동으로 위치가 설정됩니다.")
        time.sleep(3)
        self.click_position = pyautogui.position()
        self.position_label.setText(f"클릭할 위치: {self.click_position}")

    def start_macro(self):
        if not self.selected_window:
            QMessageBox.warning(self, "경고", "캡처할 창을 선택해주세요.")
            return

        repeat_count = self.repeat_spinbox.value()
        wait_time = self.time_slider.value() * 0.5
        position = self.click_position

        QMessageBox.information(self, "매크로 시작", f"3초 후 매크로가 {repeat_count}회 실행됩니다. \n선택된 창: {self.selected_window}")
        time.sleep(3)

        try:
            window = pyautogui.getWindowsWithTitle(self.selected_window)[0]
            window.activate()

            for i in range(repeat_count):
                pyautogui.hotkey('alt', 'f1')
                time.sleep(wait_time)
                pyautogui.click(position)

            QMessageBox.information(self, "매크로 완료", "매크로 실행이 완료되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"매크로 실행 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"Error in start_macro: {str(e)}")

    def select_folder(self):
        self.base_folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if self.base_folder:
            self.folder_label.setText(f"선택된 폴더: {self.base_folder}")
            self.image_folder = os.path.join(self.base_folder, "Image")
            self.cropper_folder = os.path.join(self.base_folder, "Cropper")
            
            try:
                # Image 폴더 생성 및 파일 이동
                if not os.path.exists(self.image_folder):
                    os.makedirs(self.image_folder)
                self.move_files_to_image_folder()
                
                # Cropper 폴더 생성
                if not os.path.exists(self.cropper_folder):
                    os.makedirs(self.cropper_folder)
                
                self.load_first_image()
            except Exception as e:
                QMessageBox.critical(self, "오류", f"폴더 생성 또는 파일 이동 중 오류가 발생했습니다: {str(e)}")
                logging.error(f"Error in select_folder: {str(e)}")

    def move_files_to_image_folder(self):
        image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        moved_files = 0
        for filename in os.listdir(self.base_folder):
            if filename.lower().endswith(image_extensions):
                src_path = os.path.join(self.base_folder, filename)
                dst_path = os.path.join(self.image_folder, filename)
                try:
                    shutil.move(src_path, dst_path)
                    moved_files += 1
                except Exception as e:
                    logging.error(f"Error moving file {filename}: {str(e)}")
        
        if moved_files <= 0:
            QMessageBox.warning(self, "파일 이동", "이동할 이미지 파일이 없습니다.")
            #QMessageBox.information(self, "파일 이동", f"{moved_files}개의 이미지 파일이 Image 폴더로 이동되었습니다.")
            

    def load_first_image(self):
        try:
            images = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            if not images:
                QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")
                return

            image_path = os.path.join(self.image_folder, images[0])
            self.current_image = QPixmap(image_path)
            self.image_widget.setPixmap(self.current_image)
            self.image_widget.setFixedSize(self.current_image.size())

            for spin in [self.left_spin, self.top_spin, self.right_spin, self.bottom_spin]:
                spin.setMaximum(max(self.current_image.width(), self.current_image.height()))
        except Exception as e:
            QMessageBox.critical(self, "오류", f"이미지 로드 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"Error in load_first_image: {str(e)}")

    def update_crop_coordinates(self, rect):
        self.crop_rect = rect
        self.left_spin.setValue(rect.left())
        self.top_spin.setValue(rect.top())
        self.right_spin.setValue(rect.right())
        self.bottom_spin.setValue(rect.bottom())
        self.coord_label.setText(f"크롭 좌표: ({rect.left()}, {rect.top()}, {rect.right()}, {rect.bottom()})")

    def update_crop_from_spinbox(self):
        left = self.left_spin.value()
        top = self.top_spin.value()
        right = self.right_spin.value()
        bottom = self.bottom_spin.value()
        self.crop_rect = QRect(left, top, right - left, bottom - top)
        self.image_widget.rubberband = self.crop_rect
        self.image_widget.update()
        self.coord_label.setText(f"크롭 좌표: ({left}, {top}, {right}, {bottom})")

    def update_zoom(self, value):
        zoom_factor = value / 100.0
        self.image_widget.set_zoom(zoom_factor)
        self.zoom_value_label.setText(f"{value}%")

    def create_pdf(self):
        if self.crop_rect.isNull():
            QMessageBox.warning(self, "경고", "크롭 영역을 선택해주세요.")
            return
        
        pdf_name = "cropped_ebook.pdf"
        pdf_path = os.path.join(self.cropper_folder, pdf_name)
        
        try:
            c = canvas.Canvas(pdf_path, pagesize=letter)
            images = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            
            for img_file in images:
                img_path = os.path.join(self.image_folder, img_file)
                img = Image.open(img_path)
                cropped_img = img.crop((self.crop_rect.left(), self.crop_rect.top(),
                                        self.crop_rect.right(), self.crop_rect.bottom()))
                cropped_img_path = os.path.join(self.cropper_folder, f"cropped_{img_file}")
                cropped_img.save(cropped_img_path)
                c.drawImage(cropped_img_path, 0, 0, width=letter[0], height=letter[1])
                c.showPage()
            
            c.save()
            QMessageBox.information(self, "완료", f"PDF 생성이 완료되었습니다.\n저장 위치: {pdf_path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"PDF 생성 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"Error in create_pdf: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IntegratedEBookApp()
    window.show()
    sys.exit(app.exec_())
