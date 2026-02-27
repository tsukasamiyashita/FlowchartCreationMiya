import sys
import os
import json
import uuid
import math
import unicodedata
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGraphicsScene, QGraphicsView, 
                             QGraphicsPathItem, QGraphicsLineItem, QGraphicsTextItem, 
                             QToolBar, QFileDialog, QInputDialog, QMessageBox,
                             QGraphicsItem, QGraphicsEllipseItem, QColorDialog, QLabel, QWidget, QStyle,
                             QDialog, QVBoxLayout, QHBoxLayout, QGroupBox, QRadioButton, QComboBox, QDoubleSpinBox, QPushButton)
from PyQt6.QtCore import Qt, QRectF, QPointF, QLineF, QMarginsF
from PyQt6.QtGui import (QPen, QBrush, QColor, QPainter, QImage, QPainterPath, 
                         QTransform, QPainterPathStroker, QAction, QActionGroup,
                         QPageSize, QPageLayout)
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewWidget, QPrinterInfo
from PyQt6.QtSvg import QSvgGenerator
import ezdxf

GRID_SIZE = 20  # グリッドのサイズ（スナップ間隔）

class FlowchartView(QGraphicsView):
    """ズーム機能と範囲選択（ラバーバンド）をサポートするカスタムビュー"""
    def __init__(self, scene):
        super().__init__(scene)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.zoom_factor = 1.15
        
        # 範囲選択（ラバーバンド）の有効化
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.setRubberBandSelectionMode(Qt.ItemSelectionMode.IntersectsItemShape)

    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0:
                self.scale(self.zoom_factor, self.zoom_factor)
            else:
                self.scale(1 / self.zoom_factor, 1 / self.zoom_factor)
        else:
            super().wheelEvent(event)


class NodeItem(QGraphicsPathItem):
    """フローチャートのノード（図形）"""
    def __init__(self, x, y, text="Node", node_type="process", node_id=None, bg_color="#E1F5FE", text_color="#000000"):
        super().__init__()
        self.node_type = node_type
        self.node_id = node_id if node_id else str(uuid.uuid4())
        self.edges = []
        
        self.bg_color = QColor(bg_color)
        self.text_color = QColor(text_color)
        self.default_pen = QPen(Qt.GlobalColor.black, 2)
        self._is_highlighted = False

        path = QPainterPath()
        if self.node_type == "process":
            path.addRect(QRectF(-50, -25, 100, 50))
        elif self.node_type == "decision":
            path.moveTo(0, -35)
            path.lineTo(60, 0)
            path.lineTo(0, 35)
            path.lineTo(-60, 0)
            path.closeSubpath()
        elif self.node_type == "data":
            path.moveTo(-35, -25)
            path.lineTo(65, -25)
            path.lineTo(35, 25)
            path.lineTo(-65, 25)
            path.closeSubpath()
        elif self.node_type == "terminal":
            path.addRoundedRect(QRectF(-50, -25, 100, 50), 25, 25)
        else:
            path.addRect(QRectF(-50, -25, 100, 50))

        self.setPath(path)
        self.setPos(x, y)
        self.setBrush(QBrush(self.bg_color))
        self.setPen(self.default_pen)
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
                      QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
                      QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        
        self.text_item = QGraphicsTextItem(text)
        self.text_item.setParentItem(self)
        self.text_item.setDefaultTextColor(self.text_color)
        self.set_text(text)

    def _update_text_pos(self):
        rect = self.boundingRect()
        text_rect = self.text_item.boundingRect()
        self.text_item.setPos(rect.center().x() - text_rect.width() / 2,
                              rect.center().y() - text_rect.height() / 2)

    def set_text(self, text):
        escaped_text = text.replace('\n', '<br>')
        self.text_item.setHtml(f"<div align='center'>{escaped_text}</div>")
        self._update_text_pos()

    def set_bg_color(self, color: QColor):
        self.bg_color = color
        self.setBrush(QBrush(self.bg_color))

    def set_text_color(self, color: QColor):
        self.text_color = color
        self.text_item.setDefaultTextColor(self.text_color)

    def set_highlight(self, active: bool):
        if self._is_highlighted != active:
            self._is_highlighted = active
            self.update() 

    def add_edge(self, edge):
        self.edges.append(edge)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self.scene():
            new_pos = value
            snapped_x = round(new_pos.x() / GRID_SIZE) * GRID_SIZE
            snapped_y = round(new_pos.y() / GRID_SIZE) * GRID_SIZE
            return QPointF(snapped_x, snapped_y)
            
        elif change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            for edge in self.edges:
                edge.update_position()
                
        return super().itemChange(change, value)

    def paint(self, painter, option, widget=None):
        if self._is_highlighted:
            painter.setPen(QPen(QColor("#FF5722"), 3, Qt.PenStyle.DashLine))
        elif self.isSelected():
            painter.setPen(QPen(QColor("#3B82F6"), 3))
        else:
            painter.setPen(self.default_pen)
            
        painter.setBrush(self.brush())
        painter.drawPath(self.path())

    def mouseDoubleClickEvent(self, event):
        current_text = self.text_item.toPlainText()
        new_text, ok = QInputDialog.getMultiLineText(
            None, "テキスト編集", "ノード名 (複数行入力可):", current_text
        )
        if ok:
            self.set_text(new_text)
        super().mouseDoubleClickEvent(event)


class WaypointItem(QGraphicsEllipseItem):
    """線を折り曲げるための中間制御ポイント（ウェイポイント）"""
    def __init__(self, x, y, edge):
        super().__init__(-6, -6, 12, 12)
        self.edge = edge
        self.setPos(x, y)
        self.setBrush(QBrush(QColor("#FF9800")))
        self.setPen(QPen(Qt.GlobalColor.white, 2))
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
                      QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges |
                      QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.setZValue(1)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self.scene():
            new_pos = value
            snapped_x = round(new_pos.x() / GRID_SIZE) * GRID_SIZE
            snapped_y = round(new_pos.y() / GRID_SIZE) * GRID_SIZE
            return QPointF(snapped_x, snapped_y)
            
        elif change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            self.edge.update_position()
            
        return super().itemChange(change, value)

    def paint(self, painter, option, widget=None):
        if self.isSelected():
            painter.setPen(QPen(QColor("#3B82F6"), 2))
        else:
            painter.setPen(self.pen())
        painter.setBrush(self.brush())
        painter.drawEllipse(self.rect())

    def mouseDoubleClickEvent(self, event):
        self.edge.remove_waypoint(self)
        super().mouseDoubleClickEvent(event)
        
    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        self.ungrabMouse()
        self.edge.check_waypoint_straightness(self)


class EdgeTextItem(QGraphicsTextItem):
    """エッジのラベルとして機能し、自由にドラッグ可能なテキストアイテム"""
    def __init__(self, text, edge):
        super().__init__(text)
        self.edge = edge
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable | 
                      QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
                      QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        self.setParentItem(edge)
        self.setDefaultTextColor(QColor("#333333"))
        
        self.manual_offset = None
        self._is_dragging = False

    def mousePressEvent(self, event):
        self._is_dragging = True
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        self._is_dragging = False
        super().mouseReleaseEvent(event)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange and self._is_dragging:
            base_pos = self.edge.get_auto_text_pos()
            if base_pos is not None:
                self.manual_offset = value - base_pos
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        new_text, ok = QInputDialog.getMultiLineText(
            None, "エッジのテキスト編集", "線上のテキスト (複数行入力可):", self.edge.raw_text
        )
        if ok:
            self.edge.set_text(new_text)

    def paint(self, painter, option, widget=None):
        option.state &= ~QStyle.StateFlag.State_Selected
        super().paint(painter, option, widget)
        
        if self.isSelected():
            painter.setPen(QPen(QColor("#3B82F6"), 1, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(self.boundingRect())


class EdgeItem(QGraphicsPathItem):
    """ノード間を結ぶエッジ（線）。ウェイポイントをサポート。"""
    def __init__(self, source_node, target_node, label=""):
        super().__init__()
        self.source_node = source_node
        self.target_node = target_node
        self.raw_text = label
        self.waypoints = []
        
        self._drag_start_pos = None
        self._potential_waypoint_index = -1
        
        self.default_pen = QPen(Qt.GlobalColor.black, 2)
        self.setPen(self.default_pen)
        self.setZValue(-1)
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        
        self.text_item = EdgeTextItem("", self)
        
        self._set_label_html(label)
        self.update_position()

    def boundingRect(self):
        extra = 10.0
        return super().boundingRect().adjusted(-extra, -extra, extra, extra)

    def shape(self):
        path = super().shape()
        stroker = QPainterPathStroker()
        stroker.setWidth(20)
        return stroker.createStroke(path)

    def _set_label_html(self, text):
        self.raw_text = text
        if text:
            escaped_text = text.replace('\n', '<br>')
            self.text_item.setHtml(f"<div style='font-weight: bold; font-family: sans-serif; text-align: center;'>{escaped_text}</div>")
        else:
            self.text_item.setHtml("")

    def set_text(self, text):
        self._set_label_html(text)
        self.update_position()

    def get_auto_text_pos(self):
        if not self.source_node or not self.target_node:
            return None
        pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
        
        if len(pts) < 2: return None
        
        mid_idx = len(pts) // 2
        p1 = pts[mid_idx - 1]
        p2 = pts[mid_idx]
        
        line = QLineF(p1, p2)
        center = line.center()
        rect = self.text_item.boundingRect()
        
        dx = line.dx()
        dy = line.dy()
        length = line.length()
        
        if length > 0:
            nx = dy / length
            ny = -dx / length
            offset = 15
            target_x = center.x() + nx * offset - rect.width() / 2
            target_y = center.y() + ny * offset - rect.height() / 2
            return QPointF(target_x, target_y)
        return QPointF(center.x() - rect.width() / 2, center.y() - rect.height() / 2)

    def update_position(self):
        if not self.source_node or not self.target_node:
            return
        
        self.prepareGeometryChange()
            
        pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
        
        path = QPainterPath()
        path.moveTo(pts[0])
        for p in pts[1:]:
            path.lineTo(p)
        self.setPath(path)
        
        if self.raw_text:
            base_pos = self.get_auto_text_pos()
            if base_pos is not None:
                if self.text_item.manual_offset is not None:
                    self.text_item.setPos(base_pos + self.text_item.manual_offset)
                else:
                    self.text_item.setPos(base_pos)

    def paint(self, painter, option, widget=None):
        if self.isSelected():
            painter.setPen(QPen(QColor("#3B82F6"), 3))
        else:
            painter.setPen(self.default_pen)
            
        painter.drawPath(self.path())
        
        if self.isSelected():
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QBrush(QColor("#3B82F6")))
            pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
            for i in range(len(pts) - 1):
                p1 = pts[i]
                p2 = pts[i+1]
                mid_x = (p1.x() + p2.x()) / 2
                mid_y = (p1.y() + p2.y()) / 2
                painter.drawEllipse(QPointF(mid_x, mid_y), 5.0, 5.0)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        
        # Ctrlキー押下時は標準の複数選択挙動を優先し、ウェイポイント生成をスキップ
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            return

        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.scenePos()
            pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
            
            for i in range(len(pts) - 1):
                p1 = pts[i]
                p2 = pts[i+1]
                mid_x = (p1.x() + p2.x()) / 2
                mid_y = (p1.y() + p2.y()) / 2
                
                dist = math.hypot(pos.x() - mid_x, pos.y() - mid_y)
                if dist < 30:
                    self._drag_start_pos = pos
                    self._potential_waypoint_index = i
                    event.accept()
                    return

    def mouseMoveEvent(self, event):
        if self._drag_start_pos is not None:
            if (event.scenePos() - self._drag_start_pos).manhattanLength() > 5:
                pos = event.scenePos()
                snapped_x = round(pos.x() / GRID_SIZE) * GRID_SIZE
                snapped_y = round(pos.y() / GRID_SIZE) * GRID_SIZE
                
                wp = WaypointItem(snapped_x, snapped_y, self)
                self.waypoints.insert(self._potential_waypoint_index, wp)
                
                self.scene().items_ref.append(wp)
                self.scene().addItem(wp)
                self.update_position()
                
                self._drag_start_pos = None
                self._potential_waypoint_index = -1
                
                wp.grabMouse()
                wp.setPos(snapped_x, snapped_y)
                event.accept()
                return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_start_pos = None
        self._potential_waypoint_index = -1
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        new_text, ok = QInputDialog.getMultiLineText(
            None, "エッジのテキスト編集", "線上のテキスト (複数行入力可):", self.raw_text
        )
        if ok:
            self.set_text(new_text)
        super().mouseDoubleClickEvent(event)

    def remove_waypoint(self, wp):
        if wp in self.waypoints:
            self.waypoints.remove(wp)
            self.scene().removeItem(wp)
            if wp in self.scene().items_ref:
                self.scene().items_ref.remove(wp)
            self.update_position()
            
    def check_waypoint_straightness(self, wp):
        if wp not in self.waypoints:
            return
        idx = self.waypoints.index(wp)
        
        p1 = self.source_node.scenePos() if idx == 0 else self.waypoints[idx - 1].scenePos()
        p2 = self.target_node.scenePos() if idx == len(self.waypoints) - 1 else self.waypoints[idx + 1].scenePos()
        p0 = wp.scenePos()
        
        line = QLineF(p1, p2)
        length = line.length()
        if length == 0:
            self.remove_waypoint(wp)
            return
            
        num = abs((p2.x() - p1.x()) * (p1.y() - p0.y()) - (p1.x() - p0.x()) * (p2.y() - p1.y()))
        dist = num / length
        
        dot_product = (p0.x() - p1.x()) * (p2.x() - p1.x()) + (p0.y() - p1.y()) * (p2.y() - p1.y())
        
        if dist < 15.0 and 0 <= dot_product <= length ** 2:
            self.remove_waypoint(wp)


class FlowchartScene(QGraphicsScene):
    def __init__(self, main_window):
        super().__init__(main_window)
        self.main_window = main_window
        self.source_node = None
        self.items_ref = [] 

    def mousePressEvent(self, event):
        tool = self.main_window.current_tool
        
        if tool in ["process", "decision", "data", "terminal"] and event.button() == Qt.MouseButton.LeftButton:
            pos = event.scenePos()
            snapped_x = round(pos.x() / GRID_SIZE) * GRID_SIZE
            snapped_y = round(pos.y() / GRID_SIZE) * GRID_SIZE
            
            node = NodeItem(snapped_x, snapped_y, text="Node", node_type=tool)
            self.items_ref.append(node)
            self.addItem(node)
            return

        if tool == "connect" and event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.scenePos(), QTransform())
            
            while item and not isinstance(item, NodeItem):
                item = item.parentItem()
                
            if isinstance(item, NodeItem):
                if self.source_node is None:
                    self.source_node = item
                    self.source_node.set_highlight(True)
                    self.main_window.statusBar().showMessage("エッジ接続モード: 2つ目のノードをクリックしてください")
                elif item != self.source_node:
                    edge = EdgeItem(self.source_node, item)
                    self.source_node.add_edge(edge)
                    item.add_edge(edge)
                    
                    self.items_ref.append(edge)
                    self.addItem(edge)
                    
                    self.source_node.set_highlight(False)
                    self.source_node = None
                    self.main_window.statusBar().showMessage("エッジ接続モード: 次のエッジの1つ目のノードをクリックしてください")
                return
            else:
                if self.source_node:
                    self.source_node.set_highlight(False)
                    self.source_node = None
                    self.main_window.statusBar().showMessage("エッジ接続モード: 1つ目のノードをクリックしてください")
                return
                
        super().mousePressEvent(event)

    def keyPressEvent(self, event):
        # DeleteキーまたはBackspaceキーで選択アイテムを削除する
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.main_window.delete_selected_items()
        super().keyPressEvent(event)


def clip_line_to_node(p_start: QPointF, p_end: QPointF, node: NodeItem) -> QPointF:
    """ノードの中心から次の点へ向かう線分を、ノードの境界輪郭線でクリップし、その交点を返す"""
    line = QLineF(p_start, p_end)
    polygon = node.mapToScene(node.path().toFillPolygon())
    
    best_p = p_start
    min_dist = float('inf')
    
    for i in range(polygon.count()):
        p_a = polygon.at(i)
        p_b = polygon.at((i + 1) % polygon.count())
        edge_line = QLineF(p_a, p_b)
        
        intersect_type, intersection_point = line.intersects(edge_line)
        if intersect_type == QLineF.IntersectionType.BoundedIntersection:
            dist = QLineF(p_start, intersection_point).length()
            if dist < min_dist:
                min_dist = dist
                best_p = intersection_point
                
    return best_p


class CustomPrintPreviewDialog(QDialog):
    """一般的なアプリに近い、設定とプレビューが一体化したカスタムダイアログ"""
    def __init__(self, main_window, has_selection=False):
        super().__init__(main_window)
        self.main_window = main_window
        self.setWindowTitle("印刷プレビューと設定")
        self.resize(1100, 750)
        
        self.printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        
        # 左側: 設定パネル
        settings_layout = QVBoxLayout()
        
        printer_group = QGroupBox("プリンタ")
        pr_layout = QVBoxLayout()
        self.printer_combo = QComboBox()
        self.available_printers = QPrinterInfo.availablePrinters()
        for p in self.available_printers:
            self.printer_combo.addItem(p.printerName(), p)
        
        default_printer = QPrinterInfo.defaultPrinter().printerName()
        idx = self.printer_combo.findText(default_printer)
        if idx >= 0:
            self.printer_combo.setCurrentIndex(idx)
            
        pr_layout.addWidget(self.printer_combo)
        printer_group.setLayout(pr_layout)
        settings_layout.addWidget(printer_group)
        
        paper_group = QGroupBox("用紙設定")
        pp_layout = QVBoxLayout()
        
        h_size = QHBoxLayout()
        h_size.addWidget(QLabel("サイズ:"))
        self.paper_size_combo = QComboBox()
        sizes = [
            ("A4", QPageSize.PageSizeId.A4),
            ("A3", QPageSize.PageSizeId.A3),
            ("B5", QPageSize.PageSizeId.B5),
            ("B4", QPageSize.PageSizeId.B4),
            ("Letter", QPageSize.PageSizeId.Letter),
        ]
        for name, sz_id in sizes:
            self.paper_size_combo.addItem(name, sz_id)
        h_size.addWidget(self.paper_size_combo)
        pp_layout.addLayout(h_size)
        
        h_ori = QHBoxLayout()
        h_ori.addWidget(QLabel("向き:"))
        self.ori_portrait = QRadioButton("縦")
        self.ori_landscape = QRadioButton("横")
        self.ori_portrait.setChecked(True)
        h_ori.addWidget(self.ori_portrait)
        h_ori.addWidget(self.ori_landscape)
        pp_layout.addLayout(h_ori)
        
        h_margin = QHBoxLayout()
        h_margin.addWidget(QLabel("余白(mm):"))
        self.margin_spin = QDoubleSpinBox()
        self.margin_spin.setRange(0, 100)
        self.margin_spin.setValue(10.0)
        h_margin.addWidget(self.margin_spin)
        pp_layout.addLayout(h_margin)
        
        paper_group.setLayout(pp_layout)
        settings_layout.addWidget(paper_group)
        
        range_group = QGroupBox("印刷範囲")
        range_layout = QVBoxLayout()
        self.radio_all = QRadioButton("図面全体")
        self.radio_view = QRadioButton("現在の表示範囲")
        self.radio_sel = QRadioButton("選択したアイテム")
        self.radio_all.setChecked(True)
        self.radio_sel.setEnabled(has_selection)
        
        range_layout.addWidget(self.radio_all)
        range_layout.addWidget(self.radio_view)
        range_layout.addWidget(self.radio_sel)
        range_group.setLayout(range_layout)
        settings_layout.addWidget(range_group)
        
        scale_group = QGroupBox("スケール設定")
        sc_layout = QVBoxLayout()
        
        self.radio_auto = QRadioButton("自動調整（ページに合わせる）")
        self.radio_custom = QRadioButton("倍率指定 (%)")
        self.radio_auto.setChecked(True)
        
        h_scale = QHBoxLayout()
        self.spin_scale = QDoubleSpinBox()
        self.spin_scale.setRange(10.0, 1000.0)
        self.spin_scale.setValue(100.0)
        self.spin_scale.setDecimals(1)
        self.spin_scale.setEnabled(False)
        h_scale.addWidget(self.radio_custom)
        h_scale.addWidget(self.spin_scale)
        
        sc_layout.addWidget(self.radio_auto)
        sc_layout.addLayout(h_scale)
        scale_group.setLayout(sc_layout)
        settings_layout.addWidget(scale_group)
        
        btn_layout = QVBoxLayout()
        self.btn_print = QPushButton("🖨️ 印刷を実行")
        self.btn_print.setStyleSheet("font-weight: bold; background-color: #DBEAFE; padding: 10px;")
        self.btn_cancel = QPushButton("キャンセル")
        
        btn_layout.addSpacing(20)
        btn_layout.addWidget(self.btn_print)
        btn_layout.addWidget(self.btn_cancel)
        settings_layout.addLayout(btn_layout)
        
        settings_layout.addStretch()
        
        # 右側: プレビューウィジェット
        self.preview_widget = QPrintPreviewWidget(self.printer)
        self.preview_widget.paintRequested.connect(self.handle_paint_request)
        
        main_layout = QHBoxLayout(self)
        left_panel = QWidget()
        left_panel.setLayout(settings_layout)
        left_panel.setFixedWidth(280)
        
        main_layout.addWidget(left_panel)
        main_layout.addWidget(self.preview_widget, stretch=1)
        
        self.radio_custom.toggled.connect(self.spin_scale.setEnabled)
        self.printer_combo.currentIndexChanged.connect(self.update_preview)
        self.paper_size_combo.currentIndexChanged.connect(self.update_preview)
        self.ori_portrait.toggled.connect(self.update_preview)
        self.ori_landscape.toggled.connect(self.update_preview)
        self.margin_spin.valueChanged.connect(self.update_preview)
        self.radio_all.toggled.connect(self.update_preview)
        self.radio_view.toggled.connect(self.update_preview)
        self.radio_sel.toggled.connect(self.update_preview)
        self.radio_auto.toggled.connect(self.update_preview)
        self.spin_scale.valueChanged.connect(self.update_preview)
        
        self.btn_print.clicked.connect(self.do_print)
        self.btn_cancel.clicked.connect(self.reject)
        
        self.update_preview()

    def update_printer_settings(self):
        printer_info = self.printer_combo.currentData()
        if printer_info:
            self.printer.setPrinterName(printer_info.printerName())
            
        page_size_id = self.paper_size_combo.currentData()
        self.printer.setPageSize(QPageSize(page_size_id))
        
        if self.ori_portrait.isChecked():
            self.printer.setPageOrientation(QPageLayout.Orientation.Portrait)
        else:
            self.printer.setPageOrientation(QPageLayout.Orientation.Landscape)
            
        m = self.margin_spin.value()
        self.printer.setPageMargins(QMarginsF(m, m, m, m), QPageLayout.Unit.Millimeter)

    def update_preview(self):
        self.update_printer_settings()
        self.preview_widget.updatePreview()

    def do_print(self):
        self.update_printer_settings()
        dialog = QPrintDialog(self.printer, self)
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            self.handle_paint_request(self.printer)
            self.accept()

    def handle_paint_request(self, printer):
        if self.radio_all.isChecked():
            print_rect = self.main_window.scene.itemsBoundingRect()
            selection_only = False
        elif self.radio_view.isChecked():
            print_rect = self.main_window.view.mapToScene(self.main_window.view.viewport().rect()).boundingRect()
            selection_only = False
        else:
            rect = QRectF()
            for si in self.main_window.scene.selectedItems():
                rect = rect.united(si.sceneBoundingRect())
            print_rect = rect
            selection_only = True

        if print_rect.isEmpty() or print_rect.width() == 0 or print_rect.height() == 0:
            return
            
        selected_items = self.main_window.scene.selectedItems()
        self.main_window.scene.clearSelection()
        
        rect = QRectF(print_rect)
        rect.adjust(-5, -5, 5, 5)
        
        painter = QPainter(printer)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        page_rect = printer.pageRect(QPrinter.Unit.DevicePixel)
        
        hidden_items = []
        
        def should_keep_visible(it):
            if it in selected_items: return True
            for child in it.childItems():
                if should_keep_visible(child): return True
            return False

        for item in self.main_window.scene.items():
            if not item.isVisible():
                continue
            if isinstance(item, WaypointItem):
                item.hide()
                hidden_items.append(item)
            elif selection_only and item.parentItem() is None:
                if not should_keep_visible(item):
                    item.hide()
                    hidden_items.append(item)

        if self.radio_auto.isChecked():
            self.main_window.scene.render(painter, QRectF(page_rect), rect, Qt.AspectRatioMode.KeepAspectRatio)
        else:
            scale_percent = self.spin_scale.value() / 100.0
            base_scale = printer.resolution() / self.main_window.logicalDpiX()
            scale = scale_percent * base_scale
            
            scaled_width = rect.width() * scale
            scaled_height = rect.height() * scale
            
            x_offset = page_rect.left() + max(0, (page_rect.width() - scaled_width) / 2.0)
            y_offset = page_rect.top() + max(0, (page_rect.height() - scaled_height) / 2.0)
            
            target_rect = QRectF(x_offset, y_offset, scaled_width, scaled_height)
            self.main_window.scene.render(painter, target_rect, rect, Qt.AspectRatioMode.KeepAspectRatio)
        
        for item in hidden_items: 
            item.show()

        painter.end()
        
        for item in selected_items:
            item.setSelected(True)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 保存先パスを保持する変数
        self.current_filepath = None
        self.update_window_title()
        
        self.current_tool = "select"

        self.scene = FlowchartScene(self)
        self.scene.setBackgroundBrush(QBrush(Qt.GlobalColor.white))
        self.scene.setSceneRect(-2000, -2000, 4000, 4000)

        self.view = FlowchartView(self.scene)
        self.view.centerOn(0, 0)
        self.setCentralWidget(self.view)

        self.init_menu()
        self.init_toolbar()
        self.statusBar().showMessage("準備完了: 範囲選択や複数選択（Ctrlキー+クリック）が可能です")

    def update_window_title(self):
        """現在のファイル名に応じてウィンドウのタイトルを更新する"""
        base_title = "FlowchartCreationMiya v1.2.0"
        if self.current_filepath:
            filename = os.path.basename(self.current_filepath)
            self.setWindowTitle(f"{filename} - {base_title}")
        else:
            self.setWindowTitle(base_title)

    def init_menu(self):
        menubar = self.menuBar()
        
        file_menu = menubar.addMenu("ファイル(&F)")
        
        save_action = QAction("上書き保存(&S)", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("名前を付けて保存(&A)...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)
        
        load_action = QAction("読込(&O)", self)
        load_action.setShortcut("Ctrl+O")
        load_action.triggered.connect(self.load_json)
        file_menu.addAction(load_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("エクスポート(&E)...", self)
        export_action.triggered.connect(self.export_file)
        file_menu.addAction(export_action)
        
        copy_action = QAction("Jw_cadへコピー(&C)", self)
        copy_action.triggered.connect(self.copy_to_jwcad)
        file_menu.addAction(copy_action)
        
        print_action = QAction("印刷(&P)...", self)
        print_action.triggered.connect(self.open_print_dialog)
        file_menu.addAction(print_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("終了(&X)", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 追加機能: 配置（複数選択時の整列など）
        arrange_menu = menubar.addMenu("配置(&A)")
        
        align_left_action = QAction("左に揃える", self)
        align_left_action.triggered.connect(self.align_items_left)
        arrange_menu.addAction(align_left_action)
        
        align_top_action = QAction("上に揃える", self)
        align_top_action.triggered.connect(self.align_items_top)
        arrange_menu.addAction(align_top_action)

        help_menu = menubar.addMenu("ヘルプ(&H)")
        
        usage_action = QAction("使い方(&U)", self)
        usage_action.triggered.connect(self.show_usage)
        help_menu.addAction(usage_action)
        
        about_action = QAction("バージョン情報(&A)", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def show_usage(self):
        usage_text = (
            "【ツールの操作】\n"
            "・選択モード: アイテムの移動や複数選択（ドラッグで範囲指定、Ctrl+クリックで個別選択）ができます\n"
            "・追加モード: 各図形ボタンを押した後、キャンバス上でクリックして配置\n"
            "・接続モード: 始点ノード、終点ノードの順にクリックして線を引く\n"
            "※ 追加・接続は連続で行えます。終了時は「選択」ボタンを押してください。\n\n"
            "【その他の操作】\n"
            "・一括編集: 複数選択してから「背景色」「文字色」「削除」等を実行すると一括で反映されます\n"
            "・削除: アイテムを選択してツールバーの「削除」ボタン、または Delete / Backspace キー\n"
            "・テキスト編集: ノードまたは線をダブルクリック\n"
            "・文字の移動: 線の文字をドラッグして自由に配置可能\n"
            "・ズーム: Ctrl + マウスホイール\n"
            "・スクロール: マウスホイール\n\n"
            "【ファイル操作・印刷・Jw_cad連携】\n"
            "上部のメニューバーの「ファイル」から、保存・読込・印刷などの\n"
            "すべてのアクションを実行できます。"
        )
        QMessageBox.information(self, "使い方", usage_text)

    def show_about(self):
        QMessageBox.about(self, "バージョン情報", 
                          "FlowchartCreationMiya\n"
                          "Version: v1.2.0\n\n"
                          "Python & PyQt6 製フローチャート作成ツール")

    def init_toolbar(self):
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.setStyleSheet("""
            QToolBar {
                background: #F8F9FA;
                spacing: 6px;
                padding: 6px;
                border-bottom: 1px solid #DEE2E6;
            }
            QToolButton {
                font-size: 14px;
                padding: 6px 10px;
                min-width: 55px;
                border-radius: 6px;
                background: transparent;
                color: #212529;
                text-align: center;
            }
            QToolButton:hover {
                background: #E2E6EA;
            }
            QToolButton:checked {
                background: #DBEAFE;
                color: #1D4ED8;
                font-weight: bold;
                border: 1px solid #93C5FD;
            }
            QLabel {
                font-size: 14px;
                font-weight: bold;
                color: #495057;
                padding: 0 4px;
            }
        """)

        self.action_group = QActionGroup(self)
        self.action_group.setExclusive(True)

        def add_spacer():
            spacer = QWidget()
            spacer.setFixedWidth(6)
            toolbar.addWidget(spacer)

        self.btn_select = toolbar.addAction("👆 選択")
        self.btn_select.setCheckable(True)
        self.btn_select.setChecked(True)
        self.btn_select.triggered.connect(lambda: self.set_tool("select"))
        self.action_group.addAction(self.btn_select)

        add_spacer()

        toolbar.addWidget(QLabel("➕ 追加:"))
        
        btn_process = toolbar.addAction("⬛ 処理")
        btn_process.setCheckable(True)
        btn_process.triggered.connect(lambda: self.set_tool("process"))
        self.action_group.addAction(btn_process)
        
        btn_decision = toolbar.addAction("◆ 分岐")
        btn_decision.setCheckable(True)
        btn_decision.triggered.connect(lambda: self.set_tool("decision"))
        self.action_group.addAction(btn_decision)
        
        btn_data = toolbar.addAction("▱ データ")
        btn_data.setCheckable(True)
        btn_data.triggered.connect(lambda: self.set_tool("data"))
        self.action_group.addAction(btn_data)
        
        btn_terminal = toolbar.addAction("⬭ 端子")
        btn_terminal.setCheckable(True)
        btn_terminal.triggered.connect(lambda: self.set_tool("terminal"))
        self.action_group.addAction(btn_terminal)
        
        add_spacer()
        toolbar.addSeparator()
        add_spacer()
        
        self.connect_action = toolbar.addAction("🔗 エッジ接続")
        self.connect_action.setCheckable(True)
        self.connect_action.triggered.connect(lambda: self.set_tool("connect"))
        self.action_group.addAction(self.connect_action)
        
        add_spacer()
        toolbar.addSeparator()
        add_spacer()

        toolbar.addAction("🎨 背景色", self.change_bg_color)
        toolbar.addAction("🔠 文字色", self.change_text_color)
        
        # 削除ボタンを追加
        toolbar.addAction("🗑️ 削除", self.delete_selected_items)
        
        add_spacer()
        toolbar.addSeparator()
        add_spacer()
        
        toolbar.addAction("🔍 拡大", self.zoom_in)
        toolbar.addAction("🔍 縮小", self.zoom_out)
        toolbar.addAction("🔍 100%", self.zoom_reset)

    def set_tool(self, tool_name):
        self.current_tool = tool_name
        
        if self.scene.source_node:
            self.scene.source_node.set_highlight(False)
            self.scene.source_node = None
            
        if tool_name == "select":
            self.view.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
            self.view.setCursor(Qt.CursorShape.ArrowCursor)
            self.statusBar().showMessage("準備完了: 範囲選択や複数選択（Ctrlキー+クリック）が可能です")
        elif tool_name == "connect":
            self.view.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.view.setCursor(Qt.CursorShape.CrossCursor)
            self.statusBar().showMessage("エッジ接続モード: 1つ目のノードをクリックしてください")
        else:
            self.view.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.view.setCursor(Qt.CursorShape.CrossCursor)
            self.statusBar().showMessage("ノード配置モード: キャンバスをクリックして配置します")

    def change_bg_color(self):
        selected_nodes = [item for item in self.scene.selectedItems() if isinstance(item, NodeItem)]
        if not selected_nodes:
            QMessageBox.information(self, "情報", "色を変更するノードを選択してください")
            return
            
        initial_color = selected_nodes[0].bg_color
        color = QColorDialog.getColor(initial_color, self, "背景色の選択")
        if color.isValid():
            for node in selected_nodes:
                node.set_bg_color(color)

    def change_text_color(self):
        selected_nodes = [item for item in self.scene.selectedItems() if isinstance(item, NodeItem)]
        if not selected_nodes:
            QMessageBox.information(self, "情報", "色を変更するノードを選択してください")
            return
            
        initial_color = selected_nodes[0].text_color
        color = QColorDialog.getColor(initial_color, self, "テキスト色の選択")
        if color.isValid():
            for node in selected_nodes:
                node.set_text_color(color)
                
    def align_items_left(self):
        """複数選択されたノードを一番左のノードに合わせて整列する"""
        nodes = [item for item in self.scene.selectedItems() if isinstance(item, NodeItem)]
        if len(nodes) < 2:
            return
            
        min_x = min(node.scenePos().x() for node in nodes)
        for node in nodes:
            node.setPos(min_x, node.scenePos().y())

    def align_items_top(self):
        """複数選択されたノードを一番上のノードに合わせて整列する"""
        nodes = [item for item in self.scene.selectedItems() if isinstance(item, NodeItem)]
        if len(nodes) < 2:
            return
            
        min_y = min(node.scenePos().y() for node in nodes)
        for node in nodes:
            node.setPos(node.scenePos().x(), min_y)

    def delete_selected_items(self):
        """選択されている複数のアイテム（ノード、エッジ、ウェイポイント）を安全に一括削除する"""
        selected_items = self.scene.selectedItems()
        if not selected_items:
            return

        edges_to_delete = set()
        nodes_to_delete = set()
        waypoints_to_delete = set()
        
        for item in selected_items:
            if isinstance(item, NodeItem):
                nodes_to_delete.add(item)
                # ノードが削除されるなら、接続されているエッジもすべて削除する
                for edge in item.edges:
                    edges_to_delete.add(edge)
            elif isinstance(item, EdgeItem):
                edges_to_delete.add(item)
            elif isinstance(item, WaypointItem):
                waypoints_to_delete.add(item)

        # ウェイポイント単独の削除（エッジ全体が消える場合は不要）
        for wp in waypoints_to_delete:
            if wp.edge not in edges_to_delete:
                wp.edge.remove_waypoint(wp)
                
        # エッジの完全削除（関連するウェイポイントや接続ノードからの参照も解除）
        for edge in edges_to_delete:
            for wp in edge.waypoints:
                if wp.scene():
                    self.scene.removeItem(wp)
                if wp in self.scene.items_ref:
                    self.scene.items_ref.remove(wp)
            edge.waypoints.clear()
            
            if edge in edge.source_node.edges:
                edge.source_node.edges.remove(edge)
            if edge in edge.target_node.edges:
                edge.target_node.edges.remove(edge)
                
            if edge.scene():
                self.scene.removeItem(edge)
            if edge in self.scene.items_ref:
                self.scene.items_ref.remove(edge)

        # ノードの削除
        for node in nodes_to_delete:
            if node.scene():
                self.scene.removeItem(node)
            if node in self.scene.items_ref:
                self.scene.items_ref.remove(node)
                
        self.statusBar().showMessage("選択したアイテムを削除しました", 3000)

    def zoom_in(self):
        self.view.scale(1.15, 1.15)

    def zoom_out(self):
        self.view.scale(1 / 1.15, 1 / 1.15)

    def zoom_reset(self):
        self.view.resetTransform()

    def save_file(self):
        if self.current_filepath:
            self._write_json(self.current_filepath)
            self.statusBar().showMessage(f"上書き保存しました: {self.current_filepath}", 5000)
            QMessageBox.information(self, "保存完了", f"上書き保存が完了しました。\n{self.current_filepath}")
        else:
            self.save_file_as()

    def save_file_as(self):
        initial_dir = os.path.dirname(self.current_filepath) if self.current_filepath else os.path.abspath(os.sep)
        path, _ = QFileDialog.getSaveFileName(self, "名前を付けて保存", initial_dir, "JSON Files (*.json)")
        if not path: return
        
        self.current_filepath = path
        self._write_json(self.current_filepath)
        self.update_window_title()
        self.statusBar().showMessage(f"保存しました: {self.current_filepath}", 5000)
        QMessageBox.information(self, "保存完了", f"保存が完了しました。\n{self.current_filepath}")

    def _write_json(self, path):
        data = {"nodes": [], "edges": []}
        for item in self.scene.items():
            if isinstance(item, NodeItem):
                data["nodes"].append({
                    "id": item.node_id,
                    "type": item.node_type,
                    "x": item.scenePos().x(),
                    "y": item.scenePos().y(),
                    "text": item.text_item.toPlainText(),
                    "bg_color": item.bg_color.name(),
                    "text_color": item.text_color.name()
                })
            elif isinstance(item, EdgeItem):
                offset_data = None
                if item.text_item.manual_offset is not None:
                    offset_data = {"x": item.text_item.manual_offset.x(), "y": item.text_item.manual_offset.y()}
                    
                data["edges"].append({
                    "source": item.source_node.node_id,
                    "target": item.target_node.node_id,
                    "label": item.raw_text,
                    "waypoints": [{"x": wp.scenePos().x(), "y": wp.scenePos().y()} for wp in item.waypoints],
                    "text_offset": offset_data
                })

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def load_json(self):
        initial_dir = os.path.dirname(self.current_filepath) if self.current_filepath else os.path.abspath(os.sep)
        path, _ = QFileDialog.getOpenFileName(self, "読込", initial_dir, "JSON Files (*.json)")
        if not path: return

        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.scene.clear()
        self.scene.items_ref.clear()
        node_map = {}

        for n_data in data.get("nodes", []):
            node_type = n_data.get("type", "process")
            bg_color = n_data.get("bg_color", "#E1F5FE")
            text_color = n_data.get("text_color", "#000000")
            
            node = NodeItem(n_data["x"], n_data["y"], n_data["text"], node_type, n_data["id"], bg_color, text_color)
            self.scene.items_ref.append(node)
            self.scene.addItem(node)
            node_map[node.node_id] = node

        for e_data in data.get("edges", []):
            source = node_map.get(e_data["source"])
            target = node_map.get(e_data["target"])
            label = e_data.get("label", "")
            wps_data = e_data.get("waypoints", [])
            
            if source and target:
                edge = EdgeItem(source, target, label)
                
                offset_data = e_data.get("text_offset")
                if offset_data:
                    edge.text_item.manual_offset = QPointF(offset_data["x"], offset_data["y"])
                
                for wp_d in wps_data:
                    wp = WaypointItem(wp_d["x"], wp_d["y"], edge)
                    edge.waypoints.append(wp)
                    self.scene.items_ref.append(wp)
                    self.scene.addItem(wp)
                
                source.add_edge(edge)
                target.add_edge(edge)
                self.scene.items_ref.append(edge)
                self.scene.addItem(edge)
                edge.update_position()
                
        self.current_filepath = path
        self.update_window_title()
        self.statusBar().showMessage(f"読み込みました: {self.current_filepath}", 5000)

    def export_file(self):
        initial_dir = os.path.dirname(self.current_filepath) if self.current_filepath else os.path.abspath(os.sep)
        path, filt = QFileDialog.getSaveFileName(self, "エクスポート", initial_dir, "PNG Files (*.png);;JPEG Files (*.jpeg *.jpg);;SVG Files (*.svg);;DXF Files (*.dxf)")
        if not path: return

        self.scene.clearSelection()
        rect = self.scene.itemsBoundingRect()
        
        hidden_items = []
        for item in self.scene.items():
            if isinstance(item, WaypointItem) and item.isVisible():
                item.hide()
                hidden_items.append(item)

        if rect.isEmpty():
            for item in hidden_items: item.show()
            return
        
        margin = 20
        rect.adjust(-margin, -margin, margin, margin)

        if path.endswith('.png') or path.endswith('.jpg') or path.endswith('.jpeg'):
            image = QImage(rect.size().toSize(), QImage.Format.Format_ARGB32)
            image.fill(Qt.GlobalColor.white)
            painter = QPainter(image)
            self.scene.render(painter, QRectF(image.rect()), rect)
            painter.end()
            image.save(path)

        elif path.endswith('.svg'):
            generator = QSvgGenerator()
            generator.setFileName(path)
            generator.setSize(rect.size().toSize())
            generator.setViewBox(rect)
            painter = QPainter(generator)
            self.scene.render(painter, QRectF(0, 0, rect.width(), rect.height()), rect)
            painter.end()

        elif path.endswith('.dxf'):
            self._export_dxf(path)

        for item in hidden_items:
            item.show()

        QMessageBox.information(self, "エクスポート完了", f"ファイルのエクスポートが完了しました。\n{path}")

    def _export_dxf(self, path):
        doc = ezdxf.new('R2010')
        msp = doc.modelspace()

        for item in self.scene.items():
            if isinstance(item, NodeItem):
                x = item.scenePos().x()
                y = -item.scenePos().y()
                
                if item.node_type == "process":
                    msp.add_lwpolyline([(x-50, y+25), (x+50, y+25), (x+50, y-25), (x-50, y-25)], close=True)
                elif item.node_type == "decision":
                    msp.add_lwpolyline([(x, y+35), (x+60, y), (x, y-35), (x-60, y)], close=True)
                elif item.node_type == "data":
                    msp.add_lwpolyline([(x-35, y+25), (x+65, y+25), (x+35, y-25), (x-65, y-25)], close=True)
                elif item.node_type == "terminal":
                    msp.add_lwpolyline([(x-35, y+25), (x+35, y+25), (x+50, y), (x+35, y-25), (x-35, y-25), (x-50, y)], close=True)
                else:
                    msp.add_lwpolyline([(x-50, y+25), (x+50, y+25), (x+50, y-25), (x-50, y-25)], close=True)

                text = item.text_item.toPlainText()
                if text:
                    lines = text.split('\n')
                    line_height = 15
                    start_y = y + (len(lines) - 1) * line_height / 2
                    for i, line in enumerate(lines):
                        msp.add_text(line, dxfattribs={'height': 12}).set_placement((x, start_y - i * line_height), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

            elif isinstance(item, EdgeItem):
                pts = [item.source_node.scenePos()] + [wp.scenePos() for wp in item.waypoints] + [item.target_node.scenePos()]
                
                if len(pts) >= 2:
                    pts[0] = clip_line_to_node(pts[0], pts[1], item.source_node)
                    pts[-1] = clip_line_to_node(pts[-1], pts[-2], item.target_node)
                    
                for i in range(len(pts) - 1):
                    x1, y1 = pts[i].x(), -pts[i].y()
                    x2, y2 = pts[i+1].x(), -pts[i+1].y()
                    msp.add_line((x1, y1), (x2, y2))
                
                edge_text = item.raw_text
                if edge_text:
                    base_pos = item.get_auto_text_pos()
                    if base_pos is not None:
                        if item.text_item.manual_offset is not None:
                            final_pos = base_pos + item.text_item.manual_offset
                        else:
                            final_pos = base_pos
                            
                        rect = item.text_item.boundingRect()
                        mid_x = final_pos.x() + rect.width() / 2
                        mid_y = -(final_pos.y() + rect.height() / 2)
                        
                        lines = edge_text.split('\n')
                        line_height = 12
                        start_y = mid_y + (len(lines) - 1) * line_height / 2
                        for i, line in enumerate(lines):
                            msp.add_text(line, dxfattribs={'height': 10}).set_placement((mid_x, start_y - i * line_height), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
        
        doc.saveas(path)

    def copy_to_jwcad(self):
        text_data = ["JwcTemp", "hq", "lc7", "lt1"]  
        for item in self.scene.items():
            if isinstance(item, NodeItem):
                x = item.scenePos().x()
                y = -item.scenePos().y() 
                
                def add_polygon(pts):
                    for i in range(len(pts)):
                        p1 = pts[i]
                        p2 = pts[(i + 1) % len(pts)]
                        text_data.append(f"{p1[0]} {p1[1]} {p2[0]} {p2[1]}")

                if item.node_type == "process":
                    add_polygon([(x-50, y-25), (x+50, y-25), (x+50, y+25), (x-50, y+25)])
                elif item.node_type == "decision":
                    add_polygon([(x, y-35), (x+60, y), (x, y+35), (x-60, y)])
                elif item.node_type == "data":
                    add_polygon([(x-35, y-25), (x+65, y-25), (x+35, y+25), (x-65, y+25)])
                elif item.node_type == "terminal":
                    add_polygon([(x-35, y-25), (x+35, y-25), (x+50, y), (x+35, y+25), (x-35, y+25), (x-50, y)])
                else:
                    add_polygon([(x-50, y-25), (x+50, y-25), (x+50, y+25), (x-50, y+25)])

                node_text = item.text_item.toPlainText()
                if node_text:
                    lines = node_text.split('\n')
                    line_height = 15
                    start_y = y + (len(lines) - 1) * line_height / 2
                    for i, line in enumerate(lines):
                        width_count = sum(2 if unicodedata.east_asian_width(c) in 'FWA' else 1 for c in line)
                        tw = width_count * 2.5
                        tx = x - tw
                        ty = start_y - i * line_height - 6.0
                        text_data.append(f'ch {tx} {ty} 10 0 "{line}')

            elif isinstance(item, EdgeItem):
                pts = [item.source_node.scenePos()] + [wp.scenePos() for wp in item.waypoints] + [item.target_node.scenePos()]
                
                if len(pts) >= 2:
                    pts[0] = clip_line_to_node(pts[0], pts[1], item.source_node)
                    pts[-1] = clip_line_to_node(pts[-1], pts[-2], item.target_node)
                    
                for i in range(len(pts) - 1):
                    p1 = (pts[i].x(), -pts[i].y())
                    p2 = (pts[i+1].x(), -pts[i+1].y())
                    text_data.append(f"{p1[0]} {p1[1]} {p2[0]} {p2[1]}")
                
                edge_text = item.raw_text
                if edge_text:
                    base_pos = item.get_auto_text_pos()
                    if base_pos is not None:
                        if item.text_item.manual_offset is not None:
                            final_pos = base_pos + item.text_item.manual_offset
                        else:
                            final_pos = base_pos
                            
                        rect_item = item.text_item.boundingRect()
                        mid_x = final_pos.x() + rect_item.width() / 2
                        mid_y = -(final_pos.y() + rect_item.height() / 2)
                        
                        lines = edge_text.split('\n')
                        line_height = 12
                        start_y = mid_y + (len(lines) - 1) * line_height / 2
                        for i, line in enumerate(lines):
                            width_count = sum(2 if unicodedata.east_asian_width(c) in 'FWA' else 1 for c in line)
                            tw = width_count * 2.5
                            tx = mid_x - tw
                            ty = start_y - i * line_height - 5.0
                            text_data.append(f'ch {tx} {ty} 10 0 "{line}')

        clipboard = QApplication.clipboard()
        clipboard_text = '\r\n'.join(text_data) + '\r\n'
        clipboard.setText(clipboard_text)
        QMessageBox.information(self, "完了", "Jw_cad用のデータをクリップボードにコピーしました。\nJw_cadを開いて「編集」→「貼り付け (Ctrl+V)」を実行してください。")

    def open_print_dialog(self):
        has_selection = len(self.scene.selectedItems()) > 0
        dialog = CustomPrintPreviewDialog(self, has_selection)
        dialog.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())