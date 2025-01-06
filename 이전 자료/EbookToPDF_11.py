import sys
import time
import pyautogui
import os
import shutil
import logging
import traceback
import pygetwindow as gw
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QSpinBox, QSlider, QWidget, QMessageBox,
                             QFileDialog, QScrollArea, QTabWidget, QGroupBox, QListWidget,
                             QSizePolicy, QProgressBar)
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
        self.offset = QPoint()

    def setPixmap(self, pixmap):
        self.pixmap = pixmap
        self.update()

    def paintEvent(self, event):
        if self.pixmap:
            painter = QPainter(self)
            scaled_pixmap = self.pixmap.scaled(self.pixmap.size() * self.zoom_factor, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            paint_x = max(0, (self.width() - scaled_pixmap.width()) / 2 + self.offset.x())
            paint_y = max(0, (self.height() - scaled_pixmap.height()) / 2 + self.offset.y())
            painter.drawPixmap(int(paint_x), int(paint_y), scaled_pixmap)
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

    def wheelEvent(self, event):
        if event.angleDelta().y() > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    def zoom_in(self):
        self.zoom_factor *= 1.1
        self.update()

    def zoom_out(self):
        self.zoom_factor /= 1.1
        self.update()

class IntegratedEBookApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("통합 eBook 캡처 및 PDF 생성기")
        self.setGeometry(100, 100, 1000, 800)
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
            QSpinBox, QSlider, QListWidget {
                background-color: #ffffff;
                border: 1px solid #bbb;
                padding: 5px;
                font-size: 14px;
            }
            QTabWidget::pane {
                border: 1px solid #d7d7d7;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #e1e1e1;
                padding: 10px 15px;
                font-size: 16px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                font-weight: bold;
            }
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                border: 2px solid #bbb;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        self.capture_tab = QWidget()
        self.capture_layout = QVBoxLayout(self.capture_tab)
        self.tab_widget.addTab(self.capture_tab, "화면 캡처")

        self.pdf_tab = QWidget()
        self.pdf_layout = QVBoxLayout(self.pdf_tab)
        self.tab_widget.addTab(self.pdf_tab, "PDF 생성")

        self.setup_capture_tab()
        self.setup_pdf_tab()

        self.window_titles = []
        self.selected_window = None

    def setup_capture_tab(self):
        settings_group = QGroupBox("캡처 설정")
        settings_layout = QVBoxLayout()

        repeat_layout = QHBoxLayout()
        self.repeat_label = QLabel("반복 횟수:", self)
        self.repeat_label.setStyleSheet("font-size: 16px;")
        repeat_layout.addWidget(self.repeat_label)
        
        self.repeat_spinbox = QSpinBox(self)
        self.repeat_spinbox.setRange(1, 2000)
        self.repeat_spinbox.setValue(1)
        self.repeat_spinbox.setStyleSheet("""
            QSpinBox {
                font-size: 16px;
                padding: 5px;
                border: 1px solid #bbb;
                border-radius: 3px;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 30px;
            }
        """)
        self.repeat_spinbox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.repeat_spinbox.setFixedHeight(40)
        repeat_layout.addWidget(self.repeat_spinbox)

        settings_layout.addLayout(repeat_layout)

        wait_layout = QHBoxLayout()
        self.slider_label = QLabel("대기 시간(초):", self)
        self.slider_label.setStyleSheet("font-size: 16px;")
        wait_layout.addWidget(self.slider_label)
        self.time_slider = QSlider(Qt.Horizontal, self)
        self.time_slider.setRange(1, 20)
        self.time_slider.setValue(1)
        self.time_slider.setTickPosition(QSlider.TicksBelow)
        self.time_slider.setTickInterval(1)
        wait_layout.addWidget(self.time_slider)
        self.slider_value_label = QLabel("0.5초", self)
        self.slider_value_label.setStyleSheet("font-size: 16px;")
        wait_layout.addWidget(self.slider_value_label)
        settings_layout.addLayout(wait_layout)

        self.time_slider.valueChanged.connect(self.update_slider_label)

        settings_group.setLayout(settings_layout)
        self.capture_layout.addWidget(settings_group)

        self.capture_layout.addSpacing(20)

        window_group = QGroupBox("캡처할 창 선택")
        window_layout = QVBoxLayout()

        self.refresh_windows_button = QPushButton("창 목록 새로고침", self)
        self.refresh_windows_button.clicked.connect(self.refresh_window_list)
        window_layout.addWidget(self.refresh_windows_button)

        self.window_list = QListWidget(self)
        self.window_list.setFixedHeight(150)
        self.window_list.itemClicked.connect(self.select_window_from_list)
        window_layout.addWidget(self.window_list)

        window_group.setLayout(window_layout)
        self.capture_layout.addWidget(window_group)

        self.capture_layout.addSpacing(20)

        mouse_group = QGroupBox("마우스 클릭 위치")
        mouse_layout = QVBoxLayout()

        self.set_mouse_position_button = QPushButton("마우스 클릭 위치 설정", self)
        self.set_mouse_position_button.clicked.connect(self.set_mouse_position)
        mouse_layout.addWidget(self.set_mouse_position_button)

        self.position_label = QLabel("클릭할 위치: (0, 0)", self)
        self.position_label.setStyleSheet("font-size: 16px;")
        mouse_layout.addWidget(self.position_label)

        mouse_group.setLayout(mouse_layout)
        self.capture_layout.addWidget(mouse_group)

        self.capture_layout.addSpacing(20)

        # 프로그레스 바 추가
        self.macro_progress_layout = QHBoxLayout()
        self.macro_progress_bar = QProgressBar(self)
        self.macro_progress_bar.setVisible(False)
        self.macro_progress_label = QLabel("0/0", self)
        self.macro_progress_label.setVisible(False)
        self.macro_progress_label.setStyleSheet("font-size: 20px; font-weight: bold;")  # 글자 크기와 굵기 변경
        self.macro_progress_layout.addWidget(self.macro_progress_bar)
        self.macro_progress_layout.addWidget(self.macro_progress_label)
        self.capture_layout.addLayout(self.macro_progress_layout)

        self.start_button = QPushButton("매크로 시작", self)
        self.start_button.clicked.connect(self.start_macro)
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                font-size: 18px;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.capture_layout.addWidget(self.start_button)

        self.capture_layout.addStretch()
        self.click_position = (0, 0)

    def setup_pdf_tab(self):
        folder_group = QGroupBox("폴더 선택")
        folder_layout = QVBoxLayout()

        select_layout = QHBoxLayout()
        self.select_folder_button = QPushButton("폴더 선택", self)
        self.select_folder_button.clicked.connect(self.select_folder)
        select_layout.addWidget(self.select_folder_button)
        self.folder_label = QLabel("선택된 폴더: 없음", self)
        select_layout.addWidget(self.folder_label)
        folder_layout.addLayout(select_layout)

        self.guide_label = QLabel("'C:/Users/user/Videos/Desktop' 폴더를 선택하세요.", self)
        self.guide_label.setStyleSheet("color: #666; font-style: italic; font-size: 14px;")
        folder_layout.addWidget(self.guide_label)

        folder_group.setLayout(folder_layout)
        self.pdf_layout.addWidget(folder_group)

        preview_group = QGroupBox("이미지 미리보기")
        preview_layout = QVBoxLayout()
        
        # 이미지 위젯 설정
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.image_widget = ImageWidget(self)
        self.scroll_area.setWidget(self.image_widget)
        preview_layout.addWidget(self.scroll_area)
        
        # 이전/다음 버튼 추가
        button_layout = QHBoxLayout()
        self.prev_button = QPushButton("< 이전 이미지", self)
        self.next_button = QPushButton("다음 이미지 >", self)
        self.prev_button.clicked.connect(self.show_previous_image)
        self.next_button.clicked.connect(self.show_next_image)
        button_layout.addWidget(self.prev_button)
        button_layout.addWidget(self.next_button)
        
        # 줌 컨트롤 추가
        self.zoom_in_button = QPushButton("확대", self)
        self.zoom_out_button = QPushButton("축소", self)
        self.zoom_in_button.clicked.connect(self.image_widget.zoom_in)
        self.zoom_out_button.clicked.connect(self.image_widget.zoom_out)
        button_layout.addWidget(self.zoom_in_button)
        button_layout.addWidget(self.zoom_out_button)
        
        preview_layout.addLayout(button_layout)
        
        preview_group.setLayout(preview_layout)
        self.pdf_layout.addWidget(preview_group)

        crop_group = QGroupBox("크롭 설정")
        crop_layout = QVBoxLayout()

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

        crop_group.setLayout(crop_layout)
        self.pdf_layout.addWidget(crop_group)

        # 프로그레스 바 추가
        self.pdf_progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)  # 초기에는 숨김
        self.progress_label = QLabel("0/0", self)
        self.progress_label.setVisible(False)
        self.progress_label.setStyleSheet("font-size: 20px; font-weight: bold;")  # 글자 크기와 굵기 변경
        self.pdf_progress_layout.addWidget(self.progress_bar)
        self.pdf_progress_layout.addWidget(self.progress_label)
        self.pdf_layout.addLayout(self.pdf_progress_layout)

        self.create_pdf_button = QPushButton("PDF 생성", self)
        self.create_pdf_button.clicked.connect(self.create_pdf)
        self.create_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 18px;
                padding: 15px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        self.pdf_layout.addWidget(self.create_pdf_button)

        self.base_folder = ""
        self.image_folder = ""
        self.cropper_folder = ""
        self.current_image = None
        self.crop_rect = QRect()

    def update_slider_label(self, value):
        self.slider_value_label.setText(f"{value * 0.5}초")

    def refresh_window_list(self):
        try:
            all_windows = [win.title for win in gw.getAllWindows() if win.title]
            limited_windows = all_windows[:5]

            self.window_list.clear()
            for title in limited_windows:
                self.window_list.addItem(title)
            
            if not limited_windows:
                QMessageBox.information(self, "정보", "현재 열린 창이 없습니다.")
        except Exception as e:
            error_message = f"창 목록을 새로고침하는 중 오류가 발생했습니다: {str(e)}"
            QMessageBox.critical(self, "오류", error_message)
            logging.error(error_message)

    def select_window_from_list(self, item):
        self.selected_window = item.text()
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
            window = gw.getWindowsWithTitle(self.selected_window)[0]
            window.activate()

            # 프로그레스 바 설정
            self.macro_progress_bar.setRange(0, repeat_count)
            self.macro_progress_bar.setValue(0)
            self.macro_progress_bar.setVisible(True)
            self.macro_progress_label.setVisible(True)

            for i in range(repeat_count):
                pyautogui.hotkey('alt', 'f1')
                time.sleep(wait_time)
                pyautogui.click(position)
                
                # 프로그레스 바 및 라벨 업데이트
                self.macro_progress_bar.setValue(i + 1)
                self.macro_progress_label.setText(f"{i+1}/{repeat_count}")
                QApplication.processEvents()  # UI 업데이트

            self.macro_progress_bar.setVisible(False)
            self.macro_progress_label.setVisible(False)
            QMessageBox.information(self, "매크로 완료", "매크로 실행이 완료되었습니다.")
        except Exception as e:
            self.macro_progress_bar.setVisible(False)
            self.macro_progress_label.setVisible(False)
            error_msg = f"매크로 실행 중 오류가 발생했습니다: {str(e)}\n\n"
            error_msg += traceback.format_exc()
            QMessageBox.critical(self, "오류", error_msg)
            logging.error(error_msg)

    def select_folder(self):
        self.base_folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if self.base_folder:
            self.folder_label.setText(f"선택된 폴더: {self.base_folder}")
            self.image_folder = os.path.join(self.base_folder, "Image")
            self.cropper_folder = os.path.join(self.base_folder, "Cropper")
            
            try:
                if not os.path.exists(self.image_folder):
                    os.makedirs(self.image_folder)
                self.move_files_to_image_folder()
                
                if not os.path.exists(self.cropper_folder):
                    os.makedirs(self.cropper_folder)
                
                self.load_first_image()
            except Exception as e:
                error_msg = f"폴더 생성 또는 파일 이동 중 오류가 발생했습니다: {str(e)}\n\n"
                error_msg += traceback.format_exc()
                QMessageBox.critical(self, "오류", error_msg)
                logging.error(error_msg)

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
                    logging.info(f"Moved file: {filename}")
                except Exception as e:
                    logging.error(f"Error moving file {filename}: {str(e)}")
        
        if moved_files > 0:
            QMessageBox.information(self, "파일 이동", f"{moved_files}개의 이미지 파일이 Image 폴더로 이동되었습니다.")
        else:
            QMessageBox.warning(self, "파일 이동", "이동할 이미지 파일이 없습니다.")

    def load_first_image(self):
        try:
            self.image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            if not self.image_files:
                QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")
                return

            self.current_image_index = 0
            self.load_image(self.current_image_index)
        except Exception as e:
            error_msg = f"이미지 로드 중 오류가 발생했습니다: {str(e)}\n\n"
            error_msg += traceback.format_exc()
            QMessageBox.critical(self, "오류", error_msg)
            logging.error(error_msg)

    def update_crop_coordinates(self, rect):
        scaled_rect = QRect(
            int(rect.x() / self.image_widget.zoom_factor),
            int(rect.y() / self.image_widget.zoom_factor),
            int(rect.width() / self.image_widget.zoom_factor),
            int(rect.height() / self.image_widget.zoom_factor)
        )
        self.crop_rect = scaled_rect
        self.left_spin.setValue(scaled_rect.left())
        self.top_spin.setValue(scaled_rect.top())
        self.right_spin.setValue(scaled_rect.right())
        self.bottom_spin.setValue(scaled_rect.bottom())
        self.coord_label.setText(f"크롭 좌표: ({scaled_rect.left()}, {scaled_rect.top()}, {scaled_rect.right()}, {scaled_rect.bottom()})")

    def update_crop_from_spinbox(self):
        left = self.left_spin.value()
        top = self.top_spin.value()
        right = self.right_spin.value()
        bottom = self.bottom_spin.value()
        self.crop_rect = QRect(left, top, right - left, bottom - top)
        self.image_widget.rubberband = QRect(
            int(left * self.image_widget.zoom_factor),
            int(top * self.image_widget.zoom_factor),
            int((right - left) * self.image_widget.zoom_factor),
            int((bottom - top) * self.image_widget.zoom_factor)
        )
        self.image_widget.update()
        self.coord_label.setText(f"크롭 좌표: ({left}, {top}, {right}, {bottom})")

    def show_previous_image(self):
        if not hasattr(self, 'current_image_index'):
            return
        self.current_image_index = max(0, self.current_image_index - 1)
        self.load_image(self.current_image_index)

    def show_next_image(self):
        if not hasattr(self, 'current_image_index'):
            return
        self.current_image_index = min(len(self.image_files) - 1, self.current_image_index + 1)
        self.load_image(self.current_image_index)        

    def load_image(self, index):
        if 0 <= index < len(self.image_files):
            try:
                image_path = os.path.join(self.image_folder, self.image_files[index])
                self.current_image = QPixmap(image_path)
                if self.current_image.isNull():
                    raise Exception(f"이미지를 불러올 수 없습니다: {image_path}")
                self.image_widget.setPixmap(self.current_image)
                self.image_widget.setFixedSize(self.current_image.size())

                for spin in [self.left_spin, self.top_spin, self.right_spin, self.bottom_spin]:
                    spin.setMaximum(max(self.current_image.width(), self.current_image.height()))

                # 버튼 상태 업데이트
                self.prev_button.setEnabled(index > 0)
                self.next_button.setEnabled(index < len(self.image_files) - 1)
            except Exception as e:
                error_msg = f"이미지 로드 중 오류가 발생했습니다: {str(e)}\n\n"
                error_msg += traceback.format_exc()
                QMessageBox.critical(self, "오류", error_msg)
                logging.error(error_msg)

    def create_pdf(self):
        if self.crop_rect.isNull():
            QMessageBox.warning(self, "경고", "크롭 영역을 선택해주세요.")
            return
        
        if not hasattr(self, 'cropper_folder') or not self.cropper_folder:
            QMessageBox.warning(self, "경고", "먼저 폴더를 선택해주세요.")
            return

        pdf_name = "cropped_ebook.pdf"
        pdf_path = os.path.join(self.cropper_folder, pdf_name)
        
        try:
            images = [f for f in os.listdir(self.image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
            if not images:
                QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")
                return

            total_steps = len(images)
            
            self.progress_bar.setRange(0, total_steps)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            self.progress_label.setVisible(True)
            
            c = canvas.Canvas(pdf_path, pagesize=letter)

            for i, img_file in enumerate(images):
                try:
                    # 이미지 자르기 및 PDF에 추가
                    img_path = os.path.join(self.image_folder, img_file)
                    img = Image.open(img_path)
                    cropped_img = img.crop((self.crop_rect.left(), self.crop_rect.top(),
                                            self.crop_rect.right(), self.crop_rect.bottom()))
                    cropped_img_path = os.path.join(self.cropper_folder, f"cropped_{img_file}")
                    cropped_img.save(cropped_img_path)
                    c.drawImage(cropped_img_path, 0, 0, width=letter[0], height=letter[1])
                    c.showPage()
                    
                    # 프로그레스 바 및 라벨 업데이트
                    self.progress_bar.setValue(i + 1)
                    self.progress_label.setText(f"{i+1}/{total_steps}")
                    QApplication.processEvents()
                except Exception as e:
                    logging.error(f"Error processing image {img_file}: {str(e)}")

            c.save()
            self.progress_bar.setVisible(False)
            self.progress_label.setVisible(False)
            QMessageBox.information(self, "완료", f"PDF 생성이 완료되었습니다.\n저장 위치: {pdf_path}")
        except Exception as e:
            error_msg = f"PDF 생성 중 오류가 발생했습니다: {str(e)}\n\n"
            error_msg += traceback.format_exc()
            self.progress_bar.setVisible(False)
            self.progress_label.setVisible(False)
            QMessageBox.critical(self, "오류", error_msg)
            logging.error(error_msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IntegratedEBookApp()
    window.show()
    sys.exit(app.exec_())
