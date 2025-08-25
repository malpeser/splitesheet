# Splitesheet by Halved :3
import sys
import math
import zipfile
import io
from pathlib import Path
from typing import List, Optional, Dict, Any

from PIL import Image
from PySide6.QtCore import Qt, QRectF, QPointF, QRect, QTimer, QSize
from PySide6.QtGui import (
    QPixmap,
    QImage,
    QPainter,
    QColor,
    QBrush,
    QPen,
    QFont,
    QIcon,
    QKeySequence,
    QFontDatabase,
)
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGraphicsView,
    QGraphicsScene,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsItem,
    QGraphicsSimpleTextItem,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QSpinBox,
    QLineEdit,
    QColorDialog,
    QListWidget,
    QListView,
    QListWidgetItem,
    QFormLayout,
    QDockWidget,
    QMessageBox,
    QInputDialog,
    QHBoxLayout,
    QSizePolicy,
)


# Styling & etc

QSS_DARK = r"""
QWidget { background: #0f1115; color: #e6eef3; font-size: 11pt; }

/* Buttons - base */
QPushButton {
  background: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:1, stop:0 #232730, stop:1 #1b1f24);
  color: #e6eef3;
  border-radius: 8px;
  padding: 6px 8px;
  border: 1px solid rgba(255,255,255,0.04);
  min-height: 20px;
}
QPushButton:hover { background: #2b3138; }
QPushButton:pressed { background: #181b1f; }

/* Role buttons */
QPushButton#addButton { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #294b2b, stop:1 #223a25); }
QPushButton#addButton:hover { background: #325834; }

QPushButton#delButton { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #552b2b, stop:1 #3a2020); }
QPushButton#delButton:hover { background: #663535; }

QPushButton#copyButton, QPushButton#pasteButton { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #1f3a66, stop:1 #183055); }
QPushButton#copyButton:hover, QPushButton#pasteButton:hover { background: #254675; }

QPushButton#exportButton { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #2b4b8a, stop:1 #1f3566); }
QPushButton#exportButton:hover { background: #33559b; }

/* Inputs */
QLineEdit, QSpinBox { background: #0c0e11; border: 1px solid rgba(255,255,255,0.04); padding: 6px; border-radius: 8px; min-height: 26px; }

/* Zone list */
QListWidget { background: transparent; border: none; }
QListWidget::item { border-radius: 8px; padding: 6px; margin: 3px; }
QListWidget::item:selected { background: rgba(255,255,255,0.03); outline: 1px solid rgba(255,255,255,0.04); }

/* Status */
QLabel#statusLabel { color: #9fb8d9; font-size: 9pt; }

/* Scrollbars */
QScrollBar:vertical { width:10px; background: transparent; margin: 6px 0 6px 0; }
QScrollBar::handle:vertical { background: rgba(255,255,255,0.06); border-radius: 4px; min-height: 20px; }
"""


def makerndpixmap(color: QColor, size=48, radius=8):
    pix = QPixmap(size, size)
    pix.fill(QColor(0, 0, 0, 0))
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(color)
    p.setPen(QPen(Qt.NoPen))
    rect = pix.rect().adjusted(2, 2, -2, -2)
    p.drawRoundedRect(rect, radius, radius)
    p.end()
    return pix

def makecolor(n: int):
    palette = []
    for i in range(n):
        h = int((i * 360 / max(1, n)) % 360)
        c = QColor()
        c.setHsl(h, 200, 150, 200)
        palette.append(c)
    return palette


def piltoqimg(img: Image.Image) -> QImage:
    img = img.convert("RGBA")
    data = img.tobytes("raw", "RGBA")
    qimg = QImage(data, img.width, img.height, QImage.Format.Format_RGBA8888)
    return qimg


def lumcolor(c: QColor) -> float:
    return 0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()


class GridItem(QGraphicsItem):
    def __init__(self, width, height, spacing: int = 1, color: QColor = QColor(255, 255, 255, 18)):
        super().__init__()
        self._w = width
        self._h = height
        self.spacing = spacing
        self.color = color
        self.setZValue(-500)

    def boundingRect(self) -> QRectF:
        return QRectF(0, 0, self._w, self._h)

    def paint(self, painter: QPainter, option, widget=None):
        if self.spacing <= 0:
            return
        pen = QPen(self.color)
        pen.setWidth(0)
        painter.setPen(pen)
        step = self.spacing
        x = 0
        while x <= self._w:
            painter.drawLine(x, 0, x, self._h)
            x += step
        y = 0
        while y <= self._h:
            painter.drawLine(0, y, self._w, y)
            y += step

    def set_size(self, w: int, h: int):
        self._w = w
        self._h = h
        self.prepareGeometryChange()

class FrameItem(QGraphicsRectItem):
    def __init__(self, zone: 'ZoneItem', frame_index: int, x: int, y: int, w: int, h: int):
        super().__init__(0, 0, w, h)
        self.setPos(x, y)
        self.setFlags(
            QGraphicsItem.ItemIsSelectable
            | QGraphicsItem.ItemIsMovable
            | QGraphicsItem.ItemSendsGeometryChanges
        )
        self.setAcceptHoverEvents(True)
        self.zone = zone
        self.frame_index = frame_index

        fill_color = QColor(zone.color.red(), zone.color.green(), zone.color.blue(), 110)
        self.setBrush(QBrush(fill_color))
        self.setPen(QPen(Qt.NoPen))

        self.handle_size = 4
        self.handle = QGraphicsRectItem(self.rect().width() - self.handle_size, self.rect().height() - self.handle_size,
                                        self.handle_size, self.handle_size, parent=self)
        self.handle.setBrush(QBrush(QColor(255, 255, 255, 200)))
        self.handle.setPen(QPen(Qt.NoPen))
        self.handle.setVisible(False)
        self.resizing = False

        font = QFont('Courier New', 10)
        font.setBold(True)
        self.label = QGraphicsSimpleTextItem(str(frame_index), parent=self)
        self.label.setFont(font)
        text_color = QColor(0, 0, 0) if lumcolor(zone.color) > 140 else QColor(255, 255, 255)
        self.label.setBrush(QBrush(text_color))
        self.label.setPos(2, 2)
        self.label.setFlag(QGraphicsItem.ItemIgnoresTransformations, True)

    def updateHandle(self):
        r = self.rect()
        self.handle.setRect(r.width() - self.handle_size, r.height() - self.handle_size,
                             self.handle_size, self.handle_size)

    def hoverEnterEvent(self, event):
        self.handle.setVisible(True)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        if not self.isSelected():
            self.handle.setVisible(False)
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event):
        pos = event.pos()
        r = self.rect()
        hx = r.width() - self.handle_size
        hy = r.height() - self.handle_size
        if pos.x() >= hx and pos.y() >= hy and event.button() == Qt.LeftButton:
            self.resizing = True
            self.origRect = QRectF(self.rect())
            self.origMouse = event.scenePos()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            delta = event.scenePos() - self.origMouse
            newW = max(1, int(self.origRect.width() + delta.x()))
            newH = max(1, int(self.origRect.height() + delta.y()))
            self.setRect(0, 0, newW, newH)
            self.updateHandle()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.resizing:
            self.resizing = False
            new_w = int(self.rect().width())
            new_h = int(self.rect().height())
            self.zone.update_frame_size(new_w, new_h)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemPositionChange and isinstance(value, QPointF):
            new_pos = value
            snapped = QPointF(round(new_pos.x()), round(new_pos.y()))
            return snapped
        return super().itemChange(change, value)


class ZoneItem(QGraphicsRectItem):
    def __init__(self, scene: QGraphicsScene, name: str, x: int, y: int, frame_w: int, frame_h: int,
                 rows: int, cols: int, pad_x: int, pad_y: int, color: QColor):
        super().__init__(x, y, cols * frame_w + (cols - 1) * pad_x, rows * frame_h + (rows - 1) * pad_y)
        self.scene = scene
        self.name = name
        self.setBrush(QBrush(QColor(color.red(), color.green(), color.blue(), 30)))
        self.setPen(QPen(Qt.NoPen))

        self.rows = rows
        self.cols = cols
        self.pad_x = pad_x
        self.pad_y = pad_y
        self.frame_w = frame_w
        self.frame_h = frame_h
        self.frames: List[FrameItem] = []
        self.color = color

        self.origin_marker = QGraphicsRectItem(0, 0, 3, 3)
        self.origin_marker.setBrush(QBrush(QColor(color.red(), color.green(), color.blue(), 180)))
        self.origin_marker.setPen(QPen(Qt.NoPen))
        self.origin_marker.setZValue(10)
        self.origin_marker.setVisible(False)
        scene.addItem(self.origin_marker)

        self.on_frame_size_changed = None

        self.generate_frames()

    def generate_frames(self):
        for f in list(self.frames):
            self.scene.removeItem(f)
        self.frames = []
        idx = 0
        base_x = int(self.rect().x())
        base_y = int(self.rect().y())
        for r in range(self.rows):
            for c in range(self.cols):
                fx = base_x + c * (self.frame_w + self.pad_x)
                fy = base_y + r * (self.frame_h + self.pad_y)
                f = FrameItem(self, idx, fx, fy, self.frame_w, self.frame_h)
                self.scene.addItem(f)
                self.frames.append(f)
                idx += 1
        self.update_origin_marker()

    def update_frame_size(self, new_w: int, new_h: int):
        self.frame_w = new_w
        self.frame_h = new_h
        w = self.cols * self.frame_w + max(0, (self.cols - 1)) * self.pad_x
        h = self.rows * self.frame_h + max(0, (self.rows - 1)) * self.pad_y
        r = self.rect()
        self.setRect(r.x(), r.y(), w, h)
        for f in self.frames:
            f.setRect(0, 0, new_w, new_h)
            f.updateHandle()
            f.label.setText(str(f.frame_index))
            f.setBrush(QBrush(QColor(self.color.red(), self.color.green(), self.color.blue(), 110)))
        self.update_origin_marker()
        if callable(self.on_frame_size_changed):
            self.on_frame_size_changed(new_w, new_h, self)

    def set_origin(self, x: int, y: int):
        r = self.rect()
        self.setRect(x, y, r.width(), r.height())
        self.generate_frames()

    def update_origin_marker(self):
        r = self.rect()
        ox = int(r.x())
        oy = int(r.y())
        self.origin_marker.setRect(ox, oy, 3, 3)
        self.origin_marker.setBrush(QBrush(QColor(self.color.red(), self.color.green(), self.color.blue(), 200)))
        self.origin_marker.setVisible(True)

    def bounding_box(self) -> QRect:
        r = self.rect()
        return QRect(int(r.x()), int(r.y()), int(r.width()), int(r.height()))

class ImageGraphicsView(QGraphicsView):
    def __init__(self, *args):
        super().__init__(*args)
        self.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform, False)
        self.setViewportUpdateMode(QGraphicsView.FullViewportUpdate)
        self.setDragMode(QGraphicsView.NoDrag)

        self._zoom = 1.0
        self._min_zoom = 0.25
        self._max_zoom = 64
        self._pan = False
        self._pan_start = QPointF()
        self.space_is_down = False

    def wheelEvent(self, event):
        angle = event.angleDelta().y()
        factor = 1.15 if angle > 0 else 1 / 1.15
        old_zoom = self._zoom
        new_zoom = max(self._min_zoom, min(self._max_zoom, self._zoom * factor))
        factor = new_zoom / old_zoom
        self._zoom = new_zoom
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.scale(factor, factor)
        event.accept()

    def mousePressEvent(self, event):
        pos = event.position().toPoint()
        item = self.itemAt(pos)
        if event.button() == Qt.LeftButton and (item is None or isinstance(item, QGraphicsPixmapItem)):
            self._pan = True
            self.setCursor(Qt.ClosedHandCursor)
            self._pan_start = event.position()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._pan:
            delta = event.position() - self._pan_start
            self._pan_start = event.position()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - int(delta.x()))
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - int(delta.y()))
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._pan:
            self._pan = False
            self.setCursor(Qt.ArrowCursor)
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Space:
            self.space_is_down = True
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        if event.key() == Qt.Key_Space:
            self.space_is_down = False
        super().keyReleaseEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Splitesheet")
        self.resize(1200, 800)

        self.scene = QGraphicsScene(self)
        self.view = ImageGraphicsView(self.scene)
        self.setCentralWidget(self.view)

        self.pil_image: Optional[Image.Image] = None
        self.qpixmap_item: Optional[QGraphicsPixmapItem] = None
        self.grid_item: Optional[GridItem] = None
        self.zones: List[ZoneItem] = []
        self.waiting_for_origin = False
        self.copied_zone: Optional[Dict[str, Any]] = None

        self.create_dock()
        self.setAcceptDrops(True)
        self.palette = makecolor(32)

    def show_status(self, text: str, timeout_ms: int = 2500):
        self.status_label.setText(text)
        if timeout_ms > 0:
            QTimer.singleShot(timeout_ms, lambda: self.status_label.setText(""))

    def create_dock(self):
        dock = QDockWidget("Controls", self)
        dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        w = QWidget()
        layout = QVBoxLayout()
        w.setLayout(layout)

        load_btn = QPushButton("Load Image")
        load_btn.setObjectName("basicButton")
        load_btn.clicked.connect(self.load_image)
        layout.addWidget(load_btn)

        self.sheet_input = QLineEdit()
        self.sheet_input.setPlaceholderText("Enter name for this spritesheet")
        layout.addWidget(QLabel("Sheet name:"))
        layout.addWidget(self.sheet_input)

        self.bg_line = QLineEdit("#ffffff")
        pick_btn = QPushButton("Pick color from image")
        pick_btn.setObjectName("basicButton")
        pick_btn.clicked.connect(self.pick_bg_color)
        color_btn = QPushButton("Open Color Dialog")
        color_btn.setObjectName("basicButton")
        color_btn.clicked.connect(self.open_color_dialog)
        layout.addWidget(QLabel("Background color to remove:"))
        layout.addWidget(self.bg_line)
        row = QWidget(); row_h = QHBoxLayout(); row_h.setContentsMargins(0,0,0,0); row.setLayout(row_h)
        row_h.addWidget(pick_btn)
        row_h.addWidget(color_btn)
        layout.addWidget(row)

        layout.addWidget(QLabel("Zones:"))
        self.zone_list = QListWidget()
        self.zone_list.setViewMode(QListView.IconMode)
        self.zone_list.setIconSize(QSize(20, 20))
        self.zone_list.setGridSize(QSize(90, 60))
        self.zone_list.setResizeMode(QListWidget.Adjust)
        self.zone_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.zone_list.setMinimumHeight(60)
        self.zone_list.setMaximumHeight(180)  
        self.zone_list.setMovement(QListView.Static)
        self.zone_list.setWrapping(True)
        self.zone_list.currentRowChanged.connect(self.on_zone_selected)
        layout.addWidget(self.zone_list)

        row2 = QWidget(); row2_h = QHBoxLayout(); row2_h.setContentsMargins(0,0,0,0); row2.setLayout(row2_h)
        add_zone_btn = QPushButton("Add zone")
        add_zone_btn.setObjectName("addButton")
        add_zone_btn.clicked.connect(self.add_zone)
        del_zone_btn = QPushButton("Delete zone")
        del_zone_btn.setObjectName("delButton")
        del_zone_btn.clicked.connect(self.delete_selected_zone)
        row2_h.addWidget(add_zone_btn)
        row2_h.addWidget(del_zone_btn)
        layout.addWidget(row2)

        row3 = QWidget(); row3_h = QHBoxLayout(); row3_h.setContentsMargins(0,0,0,0); row3.setLayout(row3_h)
        copy_zone_btn = QPushButton("Copy")
        copy_zone_btn.setObjectName("copyButton")
        copy_zone_btn.clicked.connect(self.copy_zone)
        paste_zone_btn = QPushButton("Paste")
        paste_zone_btn.setObjectName("pasteButton")
        paste_zone_btn.clicked.connect(self.paste_zone)
        row3_h.addWidget(copy_zone_btn)
        row3_h.addWidget(paste_zone_btn)
        layout.addWidget(row3)

        pick_origin_btn = QPushButton("Set origin by click")
        pick_origin_btn.setObjectName("basicButton")
        pick_origin_btn.clicked.connect(self.enable_origin_pick)
        layout.addWidget(pick_origin_btn)

        layout.addWidget(QLabel("Selected zone properties:"))
        form = QFormLayout()
        self.z_name = QLineEdit()
        self.z_x = QSpinBox(); self.z_x.setMaximum(100000)
        self.z_y = QSpinBox(); self.z_y.setMaximum(100000)
        self.z_w = QSpinBox(); self.z_w.setMaximum(100000)
        self.z_h = QSpinBox(); self.z_h.setMaximum(100000)
        self.z_rows = QSpinBox(); self.z_rows.setMaximum(1000)
        self.z_cols = QSpinBox(); self.z_cols.setMaximum(1000)
        self.z_pad_x = QSpinBox(); self.z_pad_x.setMaximum(1000)
        self.z_pad_y = QSpinBox(); self.z_pad_y.setMaximum(1000)
        self.z_color_btn = QPushButton("Open picker")
        self.z_color_btn.setObjectName("basicButton")
        self.z_color_btn.clicked.connect(self.set_zone_color)

        form.addRow("Name:", self.z_name)
        form.addRow("Origin X:", self.z_x)
        form.addRow("Origin Y:", self.z_y)
        form.addRow("Width:", self.z_w)
        form.addRow("Height:", self.z_h)
        form.addRow("Row:", self.z_rows)
        form.addRow("Col:", self.z_cols)
        form.addRow("Pad X:", self.z_pad_x)
        form.addRow("Pad Y:", self.z_pad_y)
        form.addRow("Color:", self.z_color_btn)
        layout.addLayout(form)

        apply_zone_btn = QPushButton("Apply changes")
        apply_zone_btn.setObjectName("basicButton")
        apply_zone_btn.clicked.connect(self.apply_zone_changes)
        layout.addWidget(apply_zone_btn)

        export_btn = QPushButton("Export frames as .zip")
        export_btn.setObjectName("exportButton")
        export_btn.clicked.connect(self.export_zip)
        layout.addWidget(export_btn)

        self.status_label = QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(self.status_label)

        layout.addStretch()
        dock.setWidget(w)
        self.addDockWidget(Qt.RightDockWidgetArea, dock)

    def load_image(self):
        fpath, _ = QFileDialog.getOpenFileName(self, "Open image", "", "Images (*.png *.bmp *.jpg *.gif *.webp)")
        if not fpath:
            self.show_status("Load cancelled", 1200)
            return
        self.open_image(Path(fpath))

    def open_image(self, path: Path):
        try:
            self.pil_image = Image.open(path).convert("RGBA")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to open image: {e}")
            return
        qimg = piltoqimg(self.pil_image)
        pix = QPixmap.fromImage(qimg)
        if self.qpixmap_item:
            try:
                self.scene.removeItem(self.qpixmap_item)
            except Exception:
                pass
        self.qpixmap_item = QGraphicsPixmapItem(pix)
        self.qpixmap_item.setTransformationMode(Qt.FastTransformation)
        self.scene.addItem(self.qpixmap_item)
        self.qpixmap_item.setZValue(-1000)
        if self.grid_item:
            self.scene.removeItem(self.grid_item)
        self.grid_item = GridItem(self.pil_image.width, self.pil_image.height, spacing=1)
        self.scene.addItem(self.grid_item)
        self.grid_item.setZValue(-500)
        self.scene.setSceneRect(QRectF(0, 0, self.pil_image.width, self.pil_image.height))
        self.view.resetTransform()
        self.show_status(f"Loaded {path.name}")

    def dragEnterEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls():
            url = mime.urls()[0]
            p = Path(url.toLocalFile())
            if p.exists():
                self.open_image(p)

    def add_zone(self):
        if not self.pil_image:
            self.show_status("Load an image first", 2000)
            return
        x, y = 0, 0
        frame_w, frame_h = 16, 16
        rows, cols, pad_x, pad_y = 1, 1, 0, 0
        name, ok = QInputDialog.getText(self, "Zone name", "Enter zone name:")
        if not ok:
            self.show_status("Add cancelled", 1000)
            return
        color = self.palette[len(self.zones) % len(self.palette)]
        z = ZoneItem(self.scene, name, x, y, frame_w, frame_h, rows, cols, pad_x, pad_y, color)
        def on_size_changed(w, h, zone):
            try:
                idx = self.zones.index(zone)
            except ValueError:
                return
            if self.zone_list.currentRow() == idx:
                self.z_w.setValue(w)
                self.z_h.setValue(h)
                self.show_status("Frame size updated", 1600)
        z.on_frame_size_changed = on_size_changed

        self.zones.append(z)
        self.scene.addItem(z)
        item = QListWidgetItem(z.name)
        pix = makerndpixmap(z.color)
        item.setIcon(QIcon(pix))
        item.setSizeHint(QSize(110, 80))
        self.zone_list.addItem(item)
        self.zone_list.setCurrentRow(self.zone_list.count() - 1)
        self.show_status("Zone added", 1000)

    def delete_selected_zone(self):
        idx = self.zone_list.currentRow()
        if idx < 0:
            self.show_status("No zone selected to delete", 1500)
            return
        z = self.zones.pop(idx)
        for f in list(z.frames):
            self.scene.removeItem(f)
        self.scene.removeItem(z.origin_marker)
        self.scene.removeItem(z)
        self.zone_list.takeItem(idx)
        self.show_status("Zone deleted", 1200)

    def enable_origin_pick(self):
        idx = self.zone_list.currentRow()
        if idx < 0:
            self.show_status("Select a zone first", 1600)
            return
        self.show_status("Click to set origin", 4000)
        old_handler = self.view.mousePressEvent

        def one_shot(e):
            pos = self.view.mapToScene(e.position().toPoint())
            x = int(round(pos.x()))
            y = int(round(pos.y()))
            z = self.zones[self.zone_list.currentRow()]
            z.set_origin(x, y)
            self.z_x.setValue(x)
            self.z_y.setValue(y)
            self.show_status(f"Zone origin set to ({x}, {y})", 2000)
            self.view.mousePressEvent = old_handler

        self.view.mousePressEvent = one_shot

    def copy_zone(self):
        idx = self.zone_list.currentRow()
        if idx < 0:
            self.show_status("Select a zone to copy", 1500)
            return
        z = self.zones[idx]
        r = z.rect()
        frames_offsets = []
        for f in z.frames:
            sp = f.scenePos()
            frames_offsets.append((int(sp.x() - r.x()), int(sp.y() - r.y())))
        self.copied_zone = {
            'name': z.name + '_copy',
            'frame_w': z.frame_w,
            'frame_h': z.frame_h,
            'rows': z.rows,
            'cols': z.cols,
            'pad_x': z.pad_x,
            'pad_y': z.pad_y,
            'color': z.color,
            'frames_offsets': frames_offsets,
            'origin': (int(r.x()), int(r.y())),
        }
        self.show_status("Zone copied", 1800)

    def paste_zone(self):
        if not self.copied_zone:
            self.show_status("No zone copied", 1400)
            return
        data = self.copied_zone
        ox, oy = data['origin']
        new_origin = (ox + 10, oy + 10)
        z = ZoneItem(self.scene, data['name'], new_origin[0], new_origin[1], data['frame_w'], data['frame_h'],
                     data['rows'], data['cols'], data['pad_x'], data['pad_y'], data['color'])
        for i, f in enumerate(z.frames):
            if i < len(data['frames_offsets']):
                offx, offy = data['frames_offsets'][i]
                f.setPos(new_origin[0] + offx, new_origin[1] + offy)
        def on_size_changed(w, h, zone):
            try:
                idx = self.zones.index(zone)
            except ValueError:
                return
            if self.zone_list.currentRow() == idx:
                self.z_w.setValue(w)
                self.z_h.setValue(h)
        z.on_frame_size_changed = on_size_changed
        self.zones.append(z)
        self.scene.addItem(z)
        item = QListWidgetItem(z.name)
        pix = makerndpixmap(z.color)
        item.setIcon(QIcon(pix))
        item.setSizeHint(QSize(110, 80))
        self.zone_list.addItem(item)
        self.zone_list.setCurrentRow(self.zone_list.count() - 1)
        self.show_status("Zone pasted", 1200)

    def on_zone_selected(self, idx: int):
        if idx < 0 or idx >= len(self.zones):
            return
        z = self.zones[idx]
        self.z_name.setText(z.name)
        r = z.rect()
        self.z_x.setValue(int(r.x()))
        self.z_y.setValue(int(r.y()))
        self.z_w.setValue(int(z.frame_w))
        self.z_h.setValue(int(z.frame_h))
        self.z_rows.setValue(z.rows)
        self.z_cols.setValue(z.cols)
        self.z_pad_x.setValue(z.pad_x)
        self.z_pad_y.setValue(z.pad_y)

    def apply_zone_changes(self):
        idx = self.zone_list.currentRow()
        if idx < 0:
            self.show_status("Select a zone first", 1300)
            return
        z = self.zones[idx]
        z.name = self.z_name.text()
        self.zone_list.currentItem().setText(z.name)
        nx = self.z_x.value()
        ny = self.z_y.value()
        z.setRect(nx, ny, z.rect().width(), z.rect().height())
        fw = self.z_w.value()
        fh = self.z_h.value()
        z.frame_w = fw
        z.frame_h = fh
        z.rows = self.z_rows.value()
        z.cols = self.z_cols.value()
        z.pad_x = self.z_pad_x.value()
        z.pad_y = self.z_pad_y.value()
        z.generate_frames()
        self.update_zone_list_icon(idx)
        self.show_status("Zone applied", 1000)

    def set_zone_color(self):
        idx = self.zone_list.currentRow()
        if idx < 0:
            self.show_status("Select a zone first", 1300)
            return
        col = QColorDialog.getColor()
        if not col.isValid():
            self.show_status("Color pick cancelled", 1000)
            return
        z = self.zones[idx]
        z.color = col
        z.setBrush(QBrush(QColor(col.red(), col.green(), col.blue(), 30)))
        z.update_frame_size(z.frame_w, z.frame_h)
        z.update_origin_marker()
        self.update_zone_list_icon(idx)
        self.show_status("Zone color updated", 1000)

    def update_zone_list_icon(self, idx: int):
        if idx < 0:
            return
        item = self.zone_list.item(idx)
        z = self.zones[idx]
        pix = makerndpixmap(z.color)
        item.setIcon(QIcon(pix))

    def pick_bg_color(self):
        self.show_status("Click on imagee to pick color", 4000)
        old_handler = self.view.mousePressEvent

        def one_shot(e):
            pos = self.view.mapToScene(e.position().toPoint())
            x = int(round(pos.x()))
            y = int(round(pos.y()))
            if self.pil_image and 0 <= x < self.pil_image.width and 0 <= y < self.pil_image.height:
                px = self.pil_image.getpixel((x, y))
                hexv = '#%02x%02x%02x' % (px[0], px[1], px[2])
                self.bg_line.setText(hexv)
                self.show_status(f"Picked color: {hexv}", 1800)
            else:
                self.show_status("Click was outside image", 1400)
            self.view.mousePressEvent = old_handler

        self.view.mousePressEvent = one_shot

    def open_color_dialog(self):
        col = QColorDialog.getColor()
        if col.isValid():
            hexv = '#%02x%02x%02x' % (col.red(), col.green(), col.blue())
            self.bg_line.setText(hexv)
            self.show_status(f"Picked color: {hexv}", 1400)

    def export_zip(self):
        if not self.pil_image:
            self.show_status("Load an image first", 1600)
            return
        sheet = self.sheet_input.text().strip()
        if not sheet:
            self.show_status("Please set a sheet export name", 1600)
            return
        folder = QFileDialog.getExistingDirectory(self, "Export to folder (a ZIP will also be created)")
        if not folder:
            self.show_status("Export cancelled", 1000)
            return
        out_zip_path = Path(folder) / f"{sheet}.zip"
        bg_hex = self.bg_line.text().strip()
        bg_rgb = None
        try:
            if bg_hex.startswith('#'):
                bg_hex = bg_hex[1:]
            if len(bg_hex) == 6:
                bg_rgb = (int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16))
        except Exception:
            bg_rgb = None

        with zipfile.ZipFile(out_zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for z in self.zones:
                frames_sorted = sorted(z.frames, key=lambda f: f.frame_index)
                for f in frames_sorted:
                    sp = f.scenePos()
                    fx = int(round(sp.x()))
                    fy = int(round(sp.y()))
                    fw = int(round(f.rect().width()))
                    fh = int(round(f.rect().height()))
                    sx = max(0, min(self.pil_image.width, fx))
                    sy = max(0, min(self.pil_image.height, fy))
                    sw = max(0, min(self.pil_image.width - sx, fw))
                    sh = max(0, min(self.pil_image.height - sy, fh))
                    if sw <= 0 or sh <= 0:
                        continue
                    crop = self.pil_image.crop((sx, sy, sx + sw, sy + sh)).convert("RGBA")
                    if bg_rgb is not None:
                        datas = crop.getdata()
                        newData = []
                        for item in datas:
                            if item[0] == bg_rgb[0] and item[1] == bg_rgb[1] and item[2] == bg_rgb[2]:
                                newData.append((255, 255, 255, 0))
                            else:
                                newData.append(item)
                        crop.putdata(newData)
                    fid = f.frame_index
                    filename = f"{sheet}_{z.name}{fid}.png"
                    buf = io.BytesIO()
                    crop.save(buf, format='PNG')
                    zf.writestr(filename, buf.getvalue())
        self.show_status(f"Exported ZIP to: {out_zip_path}", 3000)

    # keyboard handling
    def keyPressEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            if event.key() == Qt.Key_N:
                self.add_zone(); return
            if event.key() == Qt.Key_W:
                self.close(); return
            if event.key() == Qt.Key_C:
                self.copy_zone(); return
            if event.key() == Qt.Key_V:
                self.paste_zone(); return
        if event.key() == Qt.Key_Delete:
            self.delete_selected_zone(); return
        super().keyPressEvent(event)


# Entrance

if __name__ == '__main__':
    app = QApplication(sys.argv)

    families = QFontDatabase.families()
    preferred = ['Segoe UI', 'Helvetica', 'Arial']
    chosen = None
    for p in preferred:
        if p in families:
            chosen = p
            break
    if chosen:
        app.setFont(QFont(chosen, 10))
    else:
        app.setFont(QFont('', 10))

    mw = MainWindow()
    app.setStyleSheet(QSS_DARK)
    mw.show()
    sys.exit(app.exec())
