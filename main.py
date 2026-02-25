import sys
import time
import ctypes
from ctypes import wintypes
import numpy as np
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QSizeGrip, QMenu, QComboBox, QPushButton, QSizePolicy)
from PyQt6.QtCore import Qt, QPoint, QRect, QThread, pyqtSignal, QSize, QTimer
from PyQt6.QtGui import QPainter, QColor, QPen, QImage, QPixmap
import mss
from PIL import Image
import win32gui
import win32ui
import win32con

# Windows API constants
PW_RENDERFULLCONTENT = 0x00000002

class CaptureThread(QThread):
    image_received = pyqtSignal(QImage)
    
    def __init__(self, region=None, hwnd=None):
        super().__init__()
        self.region = region 
        self.hwnd = hwnd     
        self.running = True
        self.sct = mss.mss()

    def capture_window_background(self, hwnd):
        """Captures a window even if it is covered by other windows with correct stride/color handling."""
        try:
            # 1. Get exact Window Dimensions
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            
            if w <= 0 or h <= 0:
                return None

            # 2. Setup DCs
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            # 3. Create Bitmap with explicit 32-bit depth to avoid grayscale/corruption
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
            saveDC.SelectObject(saveBitMap)

            # 4. Capture using PrintWindow (works for covered windows)
            # We use a try-except here because some windows might fail PrintWindow
            try:
                result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), PW_RENDERFULLCONTENT)
            except:
                result = 0

            # Fallback to BitBlt if PrintWindow fails (result 0)
            if not result:
                saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)

            # 5. Retrieve bits and handle as NumPy array to fix Skew/Stride issues
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            
            # Create a 1D array first
            data = np.frombuffer(bmpstr, dtype='uint8')
            
            # Correct shape: Windows Bitmaps are BGRX (4 channels)
            # Re-shape to (height, width, 4)
            # We use bmpinfo['bmWidth'] and bmpinfo['bmHeight'] to be safe
            bw, bh = bmpinfo['bmWidth'], bmpinfo['bmHeight']
            
            try:
                # Some windows might have unexpected padding or sizes
                if data.size == bw * bh * 4:
                    data = data.reshape((bh, bw, 4))
                    data_rgb = data[:, :, [2, 1, 0]]
                else:
                    # Fallback to PIL if size mismatch
                    img = Image.frombuffer('RGB', (bw, bh), bmpstr, 'raw', 'BGRX', 0, 1)
                    data_rgb = np.array(img)
            except Exception as e:
                print(f"Reshape error: {e}")
                return None

            # 6. Cleanup
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            # 7. Create QImage
            # Ensuring it's a contiguous array for QImage
            data_rgb = np.ascontiguousarray(data_rgb)
            height, width, channel = data_rgb.shape
            bytesPerLine = 3 * width
            
            return QImage(data_rgb.data, width, height, bytesPerLine, QImage.Format.Format_RGB888)
            
        except Exception as e:
            print(f"Enhanced Capture Error: {e}")
            return None

    def run(self):
        while self.running:
            try:
                qimage = None
                
                if self.hwnd:
                    if not win32gui.IsWindow(self.hwnd):
                        self.running = False
                        break
                    
                    qimage = self.capture_window_background(self.hwnd)
                
                elif self.region:
                    # Static region mode
                    screenshot = self.sct.grab(self.region)
                    img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                    data = img.tobytes("raw", "RGB")
                    qimage = QImage(data, img.size[0], img.size[1], QImage.Format.Format_RGB888)
                
                if qimage:
                    # Crucial: Emit a COPY of the image because the buffer might be reused/deleted in next loop
                    self.image_received.emit(qimage.copy())
                
                time.sleep(0.04) 
            except Exception as e:
                print(f"Run Error: {e}")
                time.sleep(1)

    def stop(self):
        self.running = False
        self.wait()

class SelectionWindow(QWidget):
    region_selected = pyqtSignal(dict)
    
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setWindowState(Qt.WindowState.WindowFullScreen)
        self.setCursor(Qt.CursorShape.CrossCursor)
        
        self.begin = QPoint()
        self.end = QPoint()
        self.is_selecting = False

    def paintEvent(self, event):
        qp = QPainter(self)
        qp.fillRect(self.rect(), QColor(0, 0, 0, 100))
        
        if self.is_selecting:
            rect = QRect(self.begin, self.end).normalized()
            qp.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
            qp.fillRect(rect, Qt.GlobalColor.transparent)
            qp.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
            
            pen = QPen(Qt.GlobalColor.red, 2, Qt.PenStyle.SolidLine)
            qp.setPen(pen)
            qp.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.begin = event.pos()
            self.end = event.pos()
            self.is_selecting = True
            self.update()

    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.is_selecting = False
            rect = QRect(self.begin, self.end).normalized()
            
            if rect.width() > 10 and rect.height() > 10:
                region = {
                    "top": rect.y(),
                    "left": rect.x(),
                    "width": rect.width(),
                    "height": rect.height()
                }
                self.region_selected.emit(region)
                self.close()
            else:
                self.update()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()

class WindowSelector(QWidget):
    selected = pyqtSignal(int, str)
    manual_mode = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Select Window to Monitor")
        self.setFixedWidth(400)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        
        layout = QVBoxLayout(self)
        
        label = QLabel("เลือกโปรแกรมที่ต้องการแคป (รองรับทั้งเปิดบังอยู่):")
        layout.addWidget(label)
        
        self.combo = QComboBox()
        self.refresh_windows()
        layout.addWidget(self.combo)
        
        btn_layout = QHBoxLayout()
        
        self.btn_select = QPushButton("Start Capture")
        self.btn_select.clicked.connect(self.on_select)
        btn_layout.addWidget(self.btn_select)
        
        self.btn_manual = QPushButton("Select Area Manually")
        self.btn_manual.clicked.connect(self.on_manual)
        btn_layout.addWidget(self.btn_manual)
        
        layout.addLayout(btn_layout)
        
        self.btn_refresh = QPushButton("Refresh List")
        self.btn_refresh.clicked.connect(self.refresh_windows)
        layout.addWidget(self.btn_refresh)

    def refresh_windows(self):
        self.combo.clear()
        self.windows = []
        
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and title != self.windowTitle() and "CaptureMonitor" not in title:
                    # Simple filter to skip common system background windows
                    if title not in ["Settings", "Microsoft Store", "Program Manager"]:
                        self.windows.append((hwnd, title))
        
        win32gui.EnumWindows(enum_handler, None)
        self.windows.sort(key=lambda x: x[1].lower())
        for hwnd, title in self.windows:
            self.combo.addItem(title, hwnd)

    def on_select(self):
        hwnd = self.combo.currentData()
        title = self.combo.currentText()
        if hwnd:
            self.selected.emit(hwnd, title)
            self.close()

    def on_manual(self):
        self.manual_mode.emit()
        self.close()

class PiPWindow(QWidget):
    def __init__(self, controller, region=None, hwnd=None):
        super().__init__()
        self.region = region
        self.hwnd = hwnd
        self.controller = controller
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Display Label with size ignoring policy to prevent the expansion bug
        self.display_label = QLabel("Loading...")
        self.display_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.display_label.setStyleSheet("background-color: black; border: 2px solid #555;")
        self.display_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.display_label.setMinimumSize(10, 10)
        self.layout.addWidget(self.display_label)
        
        self.sizegrip = QSizeGrip(self)
        
        # Initial size
        if hwnd:
            try:
                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                # Default to 40% size
                self.resize(int(w * 0.4), int(h * 0.4))
            except:
                self.resize(400, 300)
        else:
            self.resize(region['width'] // 2, region['height'] // 2)
        
        self.old_pos = None
        
        self.capture_thread = CaptureThread(region=region, hwnd=hwnd)
        self.capture_thread.image_received.connect(self.update_image)
        self.capture_thread.start()

    def update_image(self, qimage):
        if qimage.isNull():
            return
            
        pixmap = QPixmap.fromImage(qimage)
        # Smoothly scale to the window size while keeping aspect ratio
        scaled_pixmap = pixmap.scaled(self.display_label.size(), 
                                      Qt.AspectRatioMode.KeepAspectRatio, 
                                      Qt.TransformationMode.SmoothTransformation)
        self.display_label.setPixmap(scaled_pixmap)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.old_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = event.globalPosition().toPoint() - self.old_pos
            self.move(self.pos() + delta)
            self.old_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.old_pos = None

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        reset_action = menu.addAction("Reset / Change Source (R)")
        menu.addSeparator()
        close_action = menu.addAction("Exit (Esc)")
        
        action = menu.exec(event.globalPos())
        if action == reset_action:
            self.restart_selection()
        elif action == close_action:
            self.close()

    def restart_selection(self):
        self.close()
        self.controller.start_ui()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        elif event.key() == Qt.Key.Key_R:
            self.restart_selection()

    def resizeEvent(self, event):
        # Update resizer position
        self.sizegrip.move(self.width() - self.sizegrip.width(), 
                          self.height() - self.sizegrip.height())
        super().resizeEvent(event)

    def closeEvent(self, event):
        self.capture_thread.stop()
        super().closeEvent(event)

class MainController:
    def __init__(self):
        self.selector = None
        self.selection_window = None
        self.pip_window = None

    def start_ui(self):
        if self.pip_window:
            self.pip_window.close()
        
        self.selector = WindowSelector()
        self.selector.selected.connect(self.show_pip_hwnd)
        self.selector.manual_mode.connect(self.start_manual_selection)
        self.selector.show()

    def start_manual_selection(self):
        self.selection_window = SelectionWindow()
        self.selection_window.region_selected.connect(self.show_pip_region)
        self.selection_window.show()

    def show_pip_hwnd(self, hwnd, title):
        self.pip_window = PiPWindow(self, hwnd=hwnd)
        self.pip_window.show()

    def show_pip_region(self, region):
        self.pip_window = PiPWindow(self, region=region)
        self.pip_window.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Enable High DPI scaling
    # In PyQt6, this is mostly automatic, but we ensure no manual overrides break it
    
    controller = MainController()
    controller.start_ui()
    
    sys.exit(app.exec())
