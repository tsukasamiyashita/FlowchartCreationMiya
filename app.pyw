import sys
import os
import json
import uuid
import math
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGraphicsScene, QGraphicsView, 
                             QGraphicsPathItem, QGraphicsLineItem, QGraphicsTextItem, 
                             QToolBar, QFileDialog, QInputDialog, QMessageBox,
                             QGraphicsItem, QGraphicsEllipseItem, QColorDialog, QLabel, QWidget)
from PyQt6.QtCore import Qt, QRectF, QPointF, QLineF
from PyQt6.QtGui import (QPen, QBrush, QColor, QPainter, QImage, QPainterPath, 
                         QTransform, QPainterPathStroker, QAction, QActionGroup)
from PyQt6.QtSvg import QSvgGenerator
import ezdxf

GRID_SIZE = 20  # グリッドのサイズ（スナップ間隔）

class FlowchartView(QGraphicsView):
    """ズーム機能（Ctrl+ホイール）をサポートするカスタムビュー"""
    def __init__(self, scene):
        super().__init__(scene)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        # ズーム時に「マウスカーソルのある位置」を中心に拡大縮小する設定
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.zoom_factor = 1.15

    def wheelEvent(self, event):
        # Ctrlキーを押しながらホイールを回したときだけズームする
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.angleDelta().y() > 0:
                self.scale(self.zoom_factor, self.zoom_factor)
            else:
                self.scale(1 / self.zoom_factor, 1 / self.zoom_factor)
        else:
            # それ以外は通常のスクロール（上下移動）
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
        
        self.text_item = QGraphicsTextItem(text, self)
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
        if active:
            highlight_pen = QPen(QColor("#FF5722"), 3, Qt.PenStyle.DashLine)
            self.setPen(highlight_pen)
        else:
            self.setPen(self.default_pen)

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

    def mouseDoubleClickEvent(self, event):
        self.edge.remove_waypoint(self)
        super().mouseDoubleClickEvent(event)
        
    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        self.ungrabMouse()
        self.edge.check_waypoint_straightness(self)


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
        
        pen = QPen(Qt.GlobalColor.black, 2)
        self.setPen(pen)
        self.setZValue(-1)
        self.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        
        self.text_item = QGraphicsTextItem(self)
        self.text_item.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.text_item.setDefaultTextColor(QColor("#333333"))
        
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

    def update_position(self):
        pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
        
        path = QPainterPath()
        path.moveTo(pts[0])
        for p in pts[1:]:
            path.lineTo(p)
        self.setPath(path)
        
        if self.raw_text:
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
                target_x = center.x() + nx * offset
                target_y = center.y() + ny * offset
                self.text_item.setPos(target_x - rect.width() / 2, target_y - rect.height() / 2)

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        if self.isSelected():
            painter.setPen(QPen(QColor("#FF9800"), 1))
            painter.setBrush(QBrush(QColor(255, 152, 0, 100)))
            pts = [self.source_node.scenePos()] + [wp.scenePos() for wp in self.waypoints] + [self.target_node.scenePos()]
            for i in range(len(pts) - 1):
                p1 = pts[i]
                p2 = pts[i+1]
                mid_x = (p1.x() + p2.x()) / 2
                mid_y = (p1.y() + p2.y()) / 2
                painter.drawEllipse(QPointF(mid_x, mid_y), 6, 6)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
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

    def mousePressEvent(self, event):
        tool = self.main_window.current_tool
        
        # モードに応じたクリック処理（ノード配置）
        if tool in ["process", "decision", "data", "terminal"] and event.button() == Qt.MouseButton.LeftButton:
            pos = event.scenePos()
            snapped_x = round(pos.x() / GRID_SIZE) * GRID_SIZE
            snapped_y = round(pos.y() / GRID_SIZE) * GRID_SIZE
            node = NodeItem(snapped_x, snapped_y, text="Node", node_type=tool)
            self.addItem(node)
            # 連続で配置できるよう、ツール状態はリセットせずイベント消費
            return

        # モードに応じたクリック処理（エッジ接続）
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
                    self.addItem(edge)
                    
                    self.source_node.set_highlight(False)
                    self.source_node = None
                    # 一回で選択に戻らず、継続してエッジを接続できるようにする
                    self.main_window.statusBar().showMessage("エッジ接続モード: 次のエッジの1つ目のノードをクリックしてください")
                return
            else:
                if self.source_node:
                    self.source_node.set_highlight(False)
                    self.source_node = None
                    self.main_window.statusBar().showMessage("エッジ接続モード: 1つ目のノードをクリックしてください")
                return
                
        # 選択モードの場合は通常のイベント処理（アイテムの選択やドラッグ）
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flowchart Editor v1.0.0")
        self.resize(1100, 700)

        self.current_tool = "select"  # 現在のツール状態の管理

        self.scene = FlowchartScene(self)
        self.scene.setBackgroundBrush(QBrush(Qt.GlobalColor.white))
        self.scene.setSceneRect(-2000, -2000, 4000, 4000)

        # 通常の QGraphicsView から、ズーム機能を持つカスタムの FlowchartView に変更
        self.view = FlowchartView(self.scene)
        self.view.centerOn(0, 0)
        self.setCentralWidget(self.view)

        self.init_menu()
        self.init_toolbar()
        self.statusBar().showMessage("準備完了: アイテムを選択・移動できます")

    def init_menu(self):
        menubar = self.menuBar()
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
            "・選択モード: アイテムの移動、テキスト編集、線の折り曲げなど\n"
            "・追加モード: 各図形ボタンを押した後、キャンバス上でクリックして配置\n"
            "・接続モード: 始点ノード、終点ノードの順にクリックして線を引く\n"
            "※ 追加・接続は連続で行えます。終了時は「選択」ボタンを押してください。\n\n"
            "【その他の操作】\n"
            "・テキスト編集: ノードまたは線をダブルクリック\n"
            "・ウェイポイント（折り曲げ点）の削除: ダブルクリック\n"
            "・ズーム: Ctrl + マウスホイール\n"
            "・スクロール: マウスホイール"
        )
        QMessageBox.information(self, "使い方", usage_text)

    def show_about(self):
        QMessageBox.about(self, "バージョン情報", 
                          "Flowchart Editor\n"
                          "Version: v1.0.0\n\n"
                          "Python & PyQt6 製フローチャート作成ツール")

    def init_toolbar(self):
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        # ボタンの排他選択（トグル）のためのグループ
        self.action_group = QActionGroup(self)
        self.action_group.setExclusive(True)

        def add_spacer():
            spacer = QWidget()
            spacer.setFixedWidth(10)
            toolbar.addWidget(spacer)

        # 👆 選択ボタンの追加
        self.btn_select = toolbar.addAction("👆 選択")
        self.btn_select.setCheckable(True)
        self.btn_select.setChecked(True)
        self.btn_select.triggered.connect(lambda: self.set_tool("select"))
        self.action_group.addAction(self.btn_select)

        add_spacer()

        label = QLabel(" ➕ 追加:")
        label.setStyleSheet("font-weight: bold; color: #333;")
        toolbar.addWidget(label)
        
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
        
        add_spacer()
        toolbar.addSeparator()
        add_spacer()
        
        # ズーム機能のボタン追加
        toolbar.addAction("🔍 拡大", self.zoom_in)
        toolbar.addAction("🔍 縮小", self.zoom_out)
        toolbar.addAction("🔍 100%", self.zoom_reset)

        add_spacer()
        toolbar.addSeparator()
        add_spacer()
        
        toolbar.addAction("💾 保存", self.save_json)
        toolbar.addAction("📂 読込", self.load_json)
        
        add_spacer()
        toolbar.addSeparator()
        add_spacer()
        
        toolbar.addAction("📤 エクスポート", self.export_file)

    def set_tool(self, tool_name):
        """ツール状態を切り替えるメソッド"""
        self.current_tool = tool_name
        
        # モード切替時にエッジ接続の途中状態をリセット
        if self.scene.source_node:
            self.scene.source_node.set_highlight(False)
            self.scene.source_node = None
            
        if tool_name == "select":
            self.view.setCursor(Qt.CursorShape.ArrowCursor)
            self.statusBar().showMessage("準備完了: アイテムを選択・移動できます")
        elif tool_name == "connect":
            self.view.setCursor(Qt.CursorShape.CrossCursor)
            self.statusBar().showMessage("エッジ接続モード: 1つ目のノードをクリックしてください")
        else:
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

    def zoom_in(self):
        self.view.scale(1.15, 1.15)

    def zoom_out(self):
        self.view.scale(1 / 1.15, 1 / 1.15)

    def zoom_reset(self):
        self.view.resetTransform()

    def save_json(self):
        # 先頭のディレクトリ（OSのルートパス等）を初期状態として指定
        initial_dir = os.path.abspath(os.sep)
        path, _ = QFileDialog.getSaveFileName(self, "保存", initial_dir, "JSON Files (*.json)")
        if not path: return

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
                data["edges"].append({
                    "source": item.source_node.node_id,
                    "target": item.target_node.node_id,
                    "label": item.raw_text,
                    "waypoints": [{"x": wp.scenePos().x(), "y": wp.scenePos().y()} for wp in item.waypoints]
                })

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def load_json(self):
        # 先頭のディレクトリ（OSのルートパス等）を初期状態として指定
        initial_dir = os.path.abspath(os.sep)
        path, _ = QFileDialog.getOpenFileName(self, "読込", initial_dir, "JSON Files (*.json)")
        if not path: return

        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.scene.clear()
        node_map = {}

        for n_data in data.get("nodes", []):
            node_type = n_data.get("type", "process")
            bg_color = n_data.get("bg_color", "#E1F5FE")
            text_color = n_data.get("text_color", "#000000")
            
            node = NodeItem(n_data["x"], n_data["y"], n_data["text"], node_type, n_data["id"], bg_color, text_color)
            self.scene.addItem(node)
            node_map[node.node_id] = node

        for e_data in data.get("edges", []):
            source = node_map.get(e_data["source"])
            target = node_map.get(e_data["target"])
            label = e_data.get("label", "")
            wps_data = e_data.get("waypoints", [])
            
            if source and target:
                edge = EdgeItem(source, target, label)
                
                for wp_d in wps_data:
                    wp = WaypointItem(wp_d["x"], wp_d["y"], edge)
                    edge.waypoints.append(wp)
                    self.scene.addItem(wp)
                
                source.add_edge(edge)
                target.add_edge(edge)
                self.scene.addItem(edge)
                edge.update_position()

    def export_file(self):
        # 先頭のディレクトリ（OSのルートパス等）を初期状態として指定
        initial_dir = os.path.abspath(os.sep)
        path, filt = QFileDialog.getSaveFileName(self, "エクスポート", initial_dir, "PNG Files (*.png);;JPEG Files (*.jpeg *.jpg);;SVG Files (*.svg);;DXF Files (*.dxf)")
        if not path: return

        self.scene.clearSelection()
        rect = self.scene.itemsBoundingRect()
        if rect.isEmpty(): return
        
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
                
                for i in range(len(pts) - 1):
                    x1, y1 = pts[i].x(), -pts[i].y()
                    x2, y2 = pts[i+1].x(), -pts[i+1].y()
                    msp.add_line((x1, y1), (x2, y2))
                
                edge_text = item.raw_text
                if edge_text:
                    mid_idx = len(pts) // 2
                    p1 = pts[mid_idx - 1]
                    p2 = pts[mid_idx]
                    x1, y1 = p1.x(), -p1.y()
                    x2, y2 = p2.x(), -p2.y()
                    
                    dx = x2 - x1
                    dy = y2 - y1
                    length = math.hypot(dx, dy)
                    if length > 0:
                        nx = dy / length
                        ny = -dx / length
                        offset = 15
                        mid_x = (x1 + x2) / 2 + nx * offset
                        mid_y = (y1 + y2) / 2 + ny * offset
                        
                        lines = edge_text.split('\n')
                        line_height = 12
                        start_y = mid_y + (len(lines) - 1) * line_height / 2
                        for i, line in enumerate(lines):
                            msp.add_text(line, dxfattribs={'height': 10}).set_placement((mid_x, start_y - i * line_height), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
        
        doc.saveas(path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())