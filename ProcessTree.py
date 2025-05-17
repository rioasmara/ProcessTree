import sys
import os
import re
import pandas as pd
import csv
from datetime import datetime, timedelta
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QComboBox, QTreeView, QTableView, QVBoxLayout,
    QHBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox, QSlider,
    QLabel, QSplitter, QMenu, QGridLayout, QScrollArea, QLineEdit, QStyledItemDelegate, QSizePolicy
)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor, QAction, QFontMetrics, QPainter
from PySide6.QtCore import Qt, QSize
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import numpy as np


class HighlightDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.search_text = ""

    def set_search_text(self, text):
        self.search_text = text.lower()

    def paint(self, painter, option, index):
        text = index.data(Qt.DisplayRole)
        if not text or not self.search_text:
            super().paint(painter, option, index)
            return

        if index.column() not in [0, 5]:
            super().paint(painter, option, index)
            return

        lower_text = text.lower()
        search_len = len(self.search_text)

        painter.save()
        painter.setFont(option.font)
        x = option.rect.x() + 2
        y = option.rect.y()
        height = option.rect.height()

        pos = 0
        while pos < len(text):
            match_pos = lower_text.find(self.search_text, pos)
            if match_pos == -1:
                remaining_text = text[pos:]
                painter.drawText(x, y, option.rect.width(), height, Qt.AlignLeft | Qt.AlignVCenter, remaining_text)
                break

            if match_pos > pos:
                before_text = text[pos:match_pos]
                metrics = QFontMetrics(option.font)
                width = metrics.horizontalAdvance(before_text)
                painter.drawText(x, y, option.rect.width(), height, Qt.AlignLeft | Qt.AlignVCenter, before_text)
                x += width

            matched_text = text[match_pos:match_pos + search_len]
            metrics = QFontMetrics(option.font)
            match_width = metrics.horizontalAdvance(matched_text)
            painter.fillRect(x, y, match_width, height, QColor(255, 182, 193))
            painter.drawText(x, y, option.rect.width(), height, Qt.AlignLeft | Qt.AlignVCenter, matched_text)
            x += match_width

            pos = match_pos + search_len

        painter.restore()


class MatplotlibWidget(QWidget):
    def __init__(self, parent=None, on_hour_selected=None):
        super().__init__(parent)
        self.on_hour_selected = on_hour_selected
        self.hours = []
        self.figure, self.ax = plt.subplots(figsize=(8, 2))
        self.canvas = FigureCanvas(self.figure)
        layout = QVBoxLayout()
        layout.addWidget(self.canvas)
        self.setLayout(layout)
        self.setMinimumHeight(200)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.canvas.mpl_connect('button_press_event', self.on_click)
        self.canvas.mpl_connect('button_press_event', self.on_double_click)

    def on_click(self, event):
        if event.dblclick:
            return
        if event.inaxes != self.ax or not self.hours:
            return

        x = event.xdata
        if x is None:
            return

        idx = int(round(x))
        if 0 <= idx < len(self.hours):
            selected_hour = self.hours[idx]
            if " " in selected_hour:
                selected_hour = selected_hour.split(" ")[1]
            if self.on_hour_selected:
                self.on_hour_selected(selected_hour)

    def on_double_click(self, event):
        if not event.dblclick:
            return
        if self.on_hour_selected:
            self.on_hour_selected(None)

    def update_plot(self, hours, counts):
        self.hours = hours
        self.ax.clear()
        self.ax.bar(hours, counts, width=0.3, align='center', color='skyblue', edgecolor='black')
        self.ax.set_ylabel('Process Count', fontsize=8)
        self.ax.set_xticks(hours)
        self.ax.set_xticklabels(hours, rotation=45, ha='right', fontsize=7)
        self.ax.tick_params(axis='y', labelsize=7)
        for i, count in enumerate(counts):
            self.ax.text(hours[i], count, str(count), ha='center', va='bottom', fontsize=7)
        self.figure.subplots_adjust(bottom=0.3, top=0.9, left=0.05, right=0.95)
        self.canvas.draw()
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.getcwd()
        plt.savefig(os.path.join(base_path, 'process_activity.png'))

    def sizeHint(self):
        return QSize(800, 200)


class ProcessTreeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Process Tree Analysis")
        self.setGeometry(100, 100, 1600, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        top_layout.setSpacing(0)

        time_filter_widget = QWidget()
        time_filter_layout = QVBoxLayout(time_filter_widget)
        time_filter_layout.setSpacing(3)
        time_filter_layout.setContentsMargins(0, 0, 0, 0)

        labels_layout = QHBoxLayout()
        labels_layout.setSpacing(3)

        start_section = QVBoxLayout()
        start_section.setSpacing(0)
        start_section.addWidget(QLabel("Start:"))
        self.start_time_label = QLabel("No time selected")
        self.start_time_label.setFixedWidth(80)
        self.start_time_label.setWordWrap(True)
        start_section.addWidget(self.start_time_label)
        labels_layout.addLayout(start_section)

        end_section = QVBoxLayout()
        end_section.setSpacing(0)
        end_section.addWidget(QLabel("End:"))
        self.end_time_label = QLabel("No time selected")
        self.end_time_label.setFixedWidth(80)
        self.end_time_label.setWordWrap(True)
        end_section.addWidget(self.end_time_label)
        labels_layout.addLayout(end_section)

        time_filter_layout.addLayout(labels_layout)

        sliders_layout = QHBoxLayout()
        sliders_layout.setSpacing(3)
        self.start_slider = QSlider(Qt.Horizontal)
        self.start_slider.setMaximumWidth(80)
        self.start_slider.setMinimum(0)
        self.start_slider.setMaximum(100)
        self.start_slider.setEnabled(False)
        self.start_slider.valueChanged.connect(self.update_time_labels)
        sliders_layout.addWidget(self.start_slider)

        self.end_slider = QSlider(Qt.Horizontal)
        self.end_slider.setMaximumWidth(80)
        self.end_slider.setMinimum(0)
        self.end_slider.setMaximum(100)
        self.end_slider.setEnabled(False)
        self.end_slider.valueChanged.connect(self.update_time_labels)
        sliders_layout.addWidget(self.end_slider)

        time_filter_layout.addLayout(sliders_layout)

        top_layout.addWidget(time_filter_widget)
        top_layout.addStretch()

        self.user_combo = QComboBox()
        self.user_combo.addItem("All Users")
        self.user_combo.currentTextChanged.connect(self.filter_tree)
        top_layout.addWidget(self.user_combo)

        top_layout.addSpacing(5)

        self.browse_button = QPushButton("Browse CSV File")
        self.browse_button.clicked.connect(self.browse_file)
        top_layout.addWidget(self.browse_button)

        top_layout.addSpacing(10)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Process Name or Command Line")
        self.search_input.setMaximumWidth(300)
        top_layout.addWidget(self.search_input)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_and_highlight)
        top_layout.addWidget(self.search_button)

        layout.addLayout(top_layout)

        self.splitter = QSplitter(Qt.Vertical)

        tree_widget = QWidget()
        tree_layout = QVBoxLayout(tree_widget)
        tree_layout.addWidget(QLabel("<b>Process Tree</b>"))
        self.histogram_widget = MatplotlibWidget(on_hour_selected=self.on_hour_selected)
        self.tree_view = QTreeView()
        self.tree_model = QStandardItemModel()
        self.tree_model.setHorizontalHeaderLabels([
            "Process Name", "Full Path", "Account Name", "PID",
            "Created Time", "Command Line"
        ])
        self.tree_view.setModel(self.tree_model)
        self.tree_delegate = HighlightDelegate(self.tree_view)
        self.tree_view.setItemDelegate(self.tree_delegate)
        self.tree_view.setSortingEnabled(True)
        self.tree_view.setColumnWidth(0, 150)
        self.tree_view.setColumnWidth(1, 400)
        self.tree_view.setColumnWidth(2, 150)
        self.tree_view.setColumnWidth(3, 100)
        self.tree_view.setColumnWidth(4, 150)
        self.tree_view.setColumnWidth(5, 500)
        self.tree_view.setWordWrap(False)
        self.tree_view.setMinimumHeight(200)
        self.tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_view.customContextMenuRequested.connect(self.show_tree_context_menu)

        self.tree_scroll_area = QScrollArea()
        self.tree_scroll_area.setWidget(self.tree_view)
        self.tree_scroll_area.setWidgetResizable(True)
        self.tree_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.tree_scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.tree_splitter = QSplitter(Qt.Vertical)
        self.tree_splitter.addWidget(self.histogram_widget)
        self.tree_splitter.addWidget(self.tree_scroll_area)
        self.tree_splitter.setStretchFactor(0, 2)
        self.tree_splitter.setStretchFactor(1, 5)
        self.tree_splitter.setSizes([200, 400])
        tree_layout.addWidget(self.tree_splitter)
        self.splitter.addWidget(tree_widget)

        stats_metrics_widget = QWidget()
        stats_metrics_layout = QHBoxLayout(stats_metrics_widget)
        stats_metrics_layout.setSpacing(0)

        stats_widget = QWidget()
        stats_layout = QVBoxLayout(stats_widget)
        stats_layout.addWidget(QLabel("<b>Process Statistics</b>"))
        self.stats_table = QTableView()
        self.stats_model = QStandardItemModel()
        self.stats_model.setHorizontalHeaderLabels(["Application", "Count"])
        self.stats_table.setModel(self.stats_model)
        self.stats_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stats_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stats_table.setColumnWidth(0, 150)
        self.stats_table.setColumnWidth(1, 50)
        self.stats_table.setMinimumHeight(100)
        self.stats_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.stats_table.customContextMenuRequested.connect(self.show_stats_context_menu)
        stats_layout.addWidget(self.stats_table)

        table_width = self.stats_table.columnWidth(0) + self.stats_table.columnWidth(1)
        min_margin = 5
        scaling_factor = 0.02
        side_margin = int(min_margin + (scaling_factor * table_width))
        stats_layout.setContentsMargins(side_margin, 2, 2, 2)
        stats_metrics_layout.addWidget(stats_widget, stretch=0, alignment=Qt.AlignLeft)

        forensic_metrics_widget = QWidget()
        forensic_metrics_layout = QGridLayout(forensic_metrics_widget)
        forensic_metrics_layout.setSpacing(3)
        forensic_metrics_layout.setContentsMargins(0, 0, 0, 0)
        forensic_metrics_layout.addWidget(QLabel("<b>Forensic Metrics</b>"), 0, 0, 1, 2)
        self.unique_users_label = QLabel("Unique Users: 0")
        self.elevated_processes_label = QLabel("Elevated Processes: 0")
        self.unique_parents_label = QLabel("Unique Parent Processes: 0")
        self.suspicious_commands_label = QLabel("Suspicious Command Lines: 0")
        self.network_commands_label = QLabel("Network-Related Commands: 0")
        self.registry_commands_label = QLabel("Registry-Related Commands: 0")
        self.unique_users_label.setFixedWidth(200)
        self.elevated_processes_label.setFixedWidth(200)
        self.unique_parents_label.setFixedWidth(200)
        self.suspicious_commands_label.setFixedWidth(200)
        self.network_commands_label.setFixedWidth(200)
        self.registry_commands_label.setFixedWidth(200)
        forensic_metrics_layout.addWidget(self.unique_users_label, 1, 0)
        forensic_metrics_layout.addWidget(self.elevated_processes_label, 1, 1)
        forensic_metrics_layout.addWidget(self.unique_parents_label, 2, 0)
        forensic_metrics_layout.addWidget(self.suspicious_commands_label, 2, 1)
        forensic_metrics_layout.addWidget(self.network_commands_label, 3, 0)
        forensic_metrics_layout.addWidget(self.registry_commands_label, 3, 1)
        stats_metrics_layout.addWidget(forensic_metrics_widget, stretch=1, alignment=Qt.AlignLeft)

        self.splitter.addWidget(stats_metrics_widget)

        self.splitter.setStretchFactor(0, 5)
        self.splitter.setStretchFactor(1, 2)
        self.splitter.setSizes([600, 200])
        layout.addWidget(self.splitter)

        self.splitter.splitterMoved.connect(self.log_splitter_sizes)

        self.processes = {}
        self.all_usernames = set()
        self.min_time = 0
        self.max_time = 0
        self.step_size = 60
        self.suspicious_patterns = ['powershell', 'cmd', 'net', 'whoami', 'curl', 'wget']
        self.network_patterns = ['curl', 'wget', 'invoke-webrequest', 'invoke-restmethod', 'net use', 'ftp']
        self.registry_patterns = ['reg add', 'reg delete', 'reg import']
        self.flagged_paths = set()
        self.selected_hour = None

        self.lotl_binaries = [
            'addinutil.exe', 'appinstaller.exe', 'aspnet_compiler.exe', 'at.exe', 'atbroker.exe',
            'bash.exe', 'bitsadmin.exe', 'certoc.exe', 'certreq.exe', 'certutil.exe',
            'cipher.exe', 'cmd.exe', 'cmdkey.exe', 'cmdl32.exe', 'cmstp.exe',
            'colorcpl.exe', 'computerdefaults.exe', 'configsecuritypolicy.exe', 'conhost.exe',
            'control.exe', 'csc.exe', 'cscript.exe', 'customshellhost.exe', 'datasvcutil.exe',
            'desktopimgdownldr.exe', 'devicecredentialdeployment.exe', 'dfsvc.exe', 'diantz.exe',
            'diskshadow.exe', 'dnscmd.exe', 'esentutl.exe', 'eventvwr.exe', 'expand.exe',
            'explorer.exe', 'extexport.exe', 'extrac32.exe', 'findstr.exe', 'finger.exe',
            'fltmc.exe', 'forfiles.exe', 'fsutil.exe', 'ftp.exe', 'gpscript.exe',
            'hh.exe', 'imewdbld.exe', 'ie4uinit.exe', 'iediagcmd.exe', 'ieexec.exe',
            'ilasm.exe', 'infdefaultinstall.exe', 'installutil.exe', 'jsc.exe', 'ldifde.exe',
            'makecab.exe', 'mavinject.exe', 'microsoft.workflow.compiler.exe', 'mmc.exe',
            'mpcmdrun.exe', 'msbuild.exe', 'msconfig.exe', 'msdt.exe', 'msedge.exe',
            'mshta.exe', 'msiexec.exe', 'netsh.exe', 'ngen.exe', 'odbcconf.exe',
            'offlinescannershell.exe', 'onedrivestandaloneupdater.exe', 'pcalua.exe', 'pcwrun.exe',
            'pktmon.exe', 'pnputil.exe', 'presentationhost.exe', 'print.exe', 'printbrm.exe',
            'provlaunch.exe', 'psr.exe', 'rasautou.exe', 'rdrleakdiag.exe', 'reg.exe',
            'regasm.exe', 'regedit.exe', 'regini.exe', 'register-cimprovider.exe', 'regsvcs.exe',
            'regsvr32.exe', 'replace.exe', 'rpcping.exe', 'rundll32.exe', 'runonce.exe',
            'runscripthelper.exe', 'sc.exe', 'schtasks.exe', 'scriptrunner.exe', 'setres.exe',
            'settingsynchost.exe', 'ssh.exe', 'stordiag.exe', 'syncappvpublishingserver.exe',
            'tar.exe', 'ttdinject.exe', 'tttracer.exe', 'unregmp2.exe', 'vbc.exe',
            'verclsid.exe', 'wab.exe', 'wbadmin.exe', 'wbemtest.exe', 'winget.exe',
            'wlrmdr.exe', 'wmic.exe', 'workfolders.exe', 'wscript.exe', 'wsreset.exe',
            'wuauclt.exe', 'xwizard.exe', 'msedge_proxy.exe', 'msedgewebview2.exe', 'wt.exe',
            'acccheckconsole.exe', 'adplus.exe', 'agentexecutor.exe', 'appcert.exe', 'appvlp.exe',
            'bginfo.exe', 'cdb.exe', 'coregen.exe', 'createdump.exe', 'csi.exe',
            'defaultpack.exe', 'devinit.exe', 'devtoolslauncher.exe', 'dnx.exe', 'dotnet.exe',
            'dsdbutil.exe', 'dtutil.exe', 'dump64.exe', 'dumpminitool.exe', 'dxcap.exe',
            'excel.exe', 'fsi.exe', 'fsianycpu.exe', 'mftrace.exe', 'microsoft.nodejstools.pressanykey.exe',
            'msaccess.exe', 'msdeploy.exe', 'msohtmed.exe', 'mspub.exe', 'msxsl.exe',
            'ntdsutil.exe', 'openconsole.exe', 'powerpnt.exe', 'procdump.exe', 'protocolhandler.exe',
            'rcsi.exe', 'remote.exe', 'sqldumper.exe', 'sqlps.exe', 'sqltoolsps.exe',
            'squirrel.exe', 'te.exe', 'teams.exe', 'testwindowremoteagent.exe', 'tracker.exe',
            'update.exe', 'vsdiagnostics.exe', 'vsiisexelauncher.exe', 'visio.exe', 'visualuiaverifynative.exe',
            'vslaunchbrowser.exe', 'vshadow.exe', 'vsjitdebugger.exe', 'wfmformat.exe', 'wfc.exe',
            'winproj.exe', 'winword.exe', 'wsl.exe', 'devtunnel.exe', 'vsls-agent.exe',
            'vstest.console.exe', 'winfile.exe', 'xsd.exe', 'powershell.exe'
        ]

    def log_splitter_sizes(self, pos, index):
        pass

    def on_hour_selected(self, hour):
        self.selected_hour = hour
        self.filter_tree()

    def show_tree_context_menu(self, position):
        index = self.tree_view.indexAt(position)
        if not index.isValid():
            return

        parent = index.parent()
        row = index.row()
        if row >= self.tree_model.rowCount(parent):
            return

        if index.column() != 0:
            index = self.tree_model.index(row, 0, parent)
        item = self.tree_model.itemFromIndex(index)
        if item is None:
            item = self.tree_model.item(row, 0) if not parent.isValid() else None
        if item is None:
            return

        process_name = item.text()
        if not process_name or process_name == "Root":
            return

        full_paths = set(data["full_path"] for data in self.processes.values()
                         if data["process_name"] == process_name and data["full_path"] != "Unknown")
        if not full_paths:
            return

        menu = QMenu(self)
        flag_menu = QMenu("Flag Process Name", self)
        unflag_menu = QMenu("Unflag Process Name", self)
        for path in sorted(full_paths):
            flag_action = QAction(path, self)
            unflag_action = QAction(path, self)
            flag_action.setEnabled(path not in self.flagged_paths)
            unflag_action.setEnabled(path in self.flagged_paths)
            flag_action.triggered.connect(lambda _, p=path: self.flag_process(p))
            unflag_action.triggered.connect(lambda _, p=path: self.unflag_process(p))
            flag_menu.addAction(flag_action)
            unflag_menu.addAction(unflag_action)
        menu.addMenu(flag_menu)
        menu.addMenu(unflag_menu)
        menu.exec(self.tree_view.viewport().mapToGlobal(position))

    def show_stats_context_menu(self, position):
        index = self.stats_table.indexAt(position)
        if not index.isValid() or index.row() == 0:
            return

        process_name = self.stats_model.item(index.row(), 0).text()
        full_paths = set(data["full_path"] for data in self.processes.values()
                         if data["process_name"] == process_name and data["full_path"] != "Unknown")
        if not full_paths:
            return

        menu = QMenu(self)
        flag_menu = QMenu("Flag Process Name", self)
        unflag_menu = QMenu("Unflag Process Name", self)
        for path in sorted(full_paths):
            flag_action = QAction(path, self)
            unflag_action = QAction(path, self)
            flag_action.setEnabled(path not in self.flagged_paths)
            unflag_action.setEnabled(path in self.flagged_paths)
            flag_action.triggered.connect(lambda _, p=path: self.flag_process(p))
            unflag_action.triggered.connect(lambda _, p=path: self.unflag_process(p))
            flag_menu.addAction(flag_action)
            unflag_menu.addAction(unflag_action)
        menu.addMenu(flag_menu)
        menu.addMenu(unflag_menu)
        menu.exec(self.stats_table.viewport().mapToGlobal(position))

    def flag_process(self, full_path):
        self.flagged_paths.add(full_path)
        self.filter_tree()

    def unflag_process(self, full_path):
        self.flagged_paths.discard(full_path)
        self.filter_tree()

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select CSV File", "", "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.load_csv_data(file_path)

    def resize_stats_table(self):
        font_metrics = QFontMetrics(self.stats_table.font())
        max_widths = [0, 0]
        for row in range(self.stats_model.rowCount()):
            for col in range(self.stats_model.columnCount()):
                item = self.stats_model.item(row, col)
                if item:
                    text = item.text()
                    text_width = font_metrics.horizontalAdvance(text)
                    max_widths[col] = max(max_widths[col], text_width)
        padding = 30
        max_widths = [width + padding for width in max_widths]
        max_widths[1] = max(max_widths[1], 70)
        self.stats_table.setColumnWidth(0, max_widths[0])
        self.stats_table.setColumnWidth(1, max_widths[1])
        total_width = sum(max_widths)
        overhead_buffer = 100
        total_width += overhead_buffer
        self.stats_table.setMinimumWidth(total_width)
        self.stats_table.setMaximumWidth(total_width)
        self.stats_table.parentWidget().setMinimumWidth(total_width + 50)
        self.stats_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        min_margin = 5
        scaling_factor = 0.02
        side_margin = int(min_margin + (scaling_factor * total_width))
        self.stats_table.parentWidget().layout().setContentsMargins(side_margin, 2, 2, 2)

    def load_csv_data(self, file_path):
        try:
            df = pd.read_csv(file_path, encoding='utf-8', header=None, quoting=csv.QUOTE_ALL)
            self.user_combo.clear()
            self.user_combo.addItem("All Users")
            self.tree_model.removeRows(0, self.tree_model.rowCount())
            self.stats_model.removeRows(0, self.stats_model.rowCount())
            self.unique_users_label.setText("Unique Users: 0")
            self.elevated_processes_label.setText("Elevated Processes: 0")
            self.unique_parents_label.setText("Unique Parent Processes: 0")
            self.suspicious_commands_label.setText("Suspicious Command Lines: 0")
            self.network_commands_label.setText("Network-Related Commands: 0")
            self.registry_commands_label.setText("Registry-Related Commands: 0")
            self.processes.clear()
            self.all_usernames.clear()
            self.flagged_paths.clear()
            self.min_time = float('inf')
            self.max_time = float('-inf')

            message_col = None
            for col in df.columns:
                if df[col].astype(str).str.contains("Creator Subject", na=False).any():
                    message_col = col
                    break

            if message_col is None:
                raise ValueError("No column contains 'Creator Subject' (Message field).")

            for idx, message in df[message_col].items():
                if pd.isna(message) or not isinstance(message, str):
                    continue
                if "EventID=4688" not in message or "Creator Subject" not in message:
                    continue

                account_name_match = re.search(
                    r"Account Name:\s+([^\n]+?)\s+Account Domain:", message, re.DOTALL
                )
                new_pid_match = re.search(
                    r"New Process ID:\s+([^\n]+?)\s+New Process Name:", message, re.DOTALL
                )
                creator_pid_match = re.search(
                    r"Creator Process ID:\s+([^\n]+?)\s+Creator Process Name:", message, re.DOTALL
                )
                process_name_match = re.search(
                    r"New Process Name:\s+([^\n]+?)\s+Token Elevation Type:", message, re.DOTALL
                )
                time_match = re.search(
                    r"TimeGenerated=(\d+)", message
                )
                cmdline_match = re.search(
                    r"Process Command Line:\s+(.+?)(?=\s+Token)", message, re.DOTALL
                )
                elevation_match = re.search(
                    r"Token Elevation Type:\s+(.+?)\s+Mandatory Label:", message, re.DOTALL
                )
                creator_name_match = re.search(
                    r"Creator Process Name:\s+([^\n]+?)\s+Process Command Line:", message, re.DOTALL
                )

                if not all([account_name_match, new_pid_match, creator_pid_match, process_name_match]):
                    continue

                account_name = account_name_match.group(1).strip()
                new_pid = new_pid_match.group(1).strip()
                creator_pid = creator_pid_match.group(1).strip()
                full_path = process_name_match.group(1).strip()
                process_name = full_path.split('\\')[-1] if full_path else "Unknown"
                time_generated = time_match.group(1) if time_match else "0"
                cmdline = cmdline_match.group(1).strip().lower() if cmdline_match else "Unknown"
                elevation_type = elevation_match.group(1).strip() if elevation_match else "Unknown"
                creator_name = creator_name_match.group(1).strip() if creator_name_match else "Unknown"

                try:
                    time_int = int(time_generated)
                    if time_int <= 0:
                        time_int = 0
                    elif time_int < 0 or time_int > 2147483647:
                        time_int = 0
                    else:
                        self.min_time = min(self.min_time, time_int)
                        self.max_time = max(self.max_time, time_int)
                except ValueError:
                    time_int = 0

                if elevation_type not in ["TokenElevationTypeDefault (1)", "TokenElevationTypeLimited (2)",
                                          "TokenElevationTypeFull (3)", "Unknown"]:
                    pass

                if account_name and account_name != "-":
                    self.all_usernames.add(account_name)
                    self.processes[new_pid] = {
                        "account_name": account_name,
                        "creator_pid": creator_pid,
                        "process_name": process_name,
                        "full_path": full_path if full_path else "Unknown",
                        "time_generated": time_int,
                        "cmdline": cmdline,
                        "elevation_type": elevation_type,
                        "creator_name": creator_name
                    }

            for pid, data in self.processes.items():
                if data["process_name"].lower() in [bin.lower() for bin in self.lotl_binaries]:
                    self.flagged_paths.add(data["full_path"])

            if self.all_usernames:
                self.user_combo.addItems(sorted(self.all_usernames))
                self.user_combo.setCurrentIndex(0)

            if self.min_time == float('inf') or self.max_time == float('-inf'):
                self.min_time = int(datetime(2025, 5, 14).timestamp())
                self.max_time = int(datetime(2025, 5, 14, 23, 59, 59).timestamp())
            range_seconds = self.max_time - self.min_time
            self.slider_max = range_seconds // self.step_size if range_seconds > 0 else 100
            self.start_slider.setMaximum(self.slider_max)
            self.end_slider.setMaximum(self.slider_max)
            self.start_slider.setValue(0)
            self.end_slider.setValue(self.slider_max)
            self.start_slider.setEnabled(True)
            self.end_slider.setEnabled(True)
            self.update_time_labels()

            self.filter_tree()

        except Exception as e:
            self.user_combo.clear()
            self.user_combo.addItem("All Users")
            self.tree_model.removeRows(0, self.tree_model.rowCount())
            self.stats_model.removeRows(0, self.stats_model.rowCount())
            self.unique_users_label.setText("Unique Users: 0")
            self.elevated_processes_label.setText("Elevated Processes: 0")
            self.unique_parents_label.setText("Unique Parent Processes: 0")
            self.suspicious_commands_label.setText("Suspicious Command Lines: 0")
            self.network_commands_label.setText("Network-Related Commands: 0")
            self.registry_commands_label.setText("Registry-Related Commands: 0")
            self.start_slider.setEnabled(False)
            self.end_slider.setEnabled(False)
            self.start_time_label.setText("No time selected")
            self.end_time_label.setText("No time selected")
            QMessageBox.critical(self, "Error", f"Failed to load CSV file: {str(e)}")

    def update_time_labels(self):
        start_value = self.start_slider.value()
        end_value = self.end_slider.value()
        start_time = self.min_time + (start_value * self.step_size)
        end_time = self.min_time + (end_value * self.step_size)
        try:
            start_str = datetime.fromtimestamp(start_time).strftime('%Y-%m-%d %H:%M:%S')
            end_str = datetime.fromtimestamp(end_time).strftime('%Y-%m-%d %H:%M:%S')
        except (OSError, ValueError, TypeError):
            start_str = "Invalid"
            end_str = "Invalid"
        self.start_time_label.setText(start_str)
        self.end_time_label.setText(end_str)
        self.filter_tree()

    def reset_time_range(self):
        self.start_slider.setValue(0)
        self.end_slider.setValue(self.slider_max)
        self.update_time_labels()

    def build_process_hist(self, filter_username=None, time_from=None, time_to=None):
        if not self.processes:
            self.histogram_widget.update_plot([], [])
            return

        children = {}
        for pid, data in self.processes.items():
            creator_pid = data["creator_pid"]
            if creator_pid not in children:
                children[creator_pid] = []
            children[creator_pid].append(pid)

        process_list = []
        for pid, process_data in self.processes.items():
            if process_data['time_generated'] <= 0:
                continue
            if process_data['time_generated'] > 2147483647:
                continue
            is_root = process_data["creator_pid"] not in self.processes
            process_type = "Root" if is_root else f"Child of {process_data['creator_pid']}"
            process_list.append({
                'pid': pid,
                'time_generated': process_data['time_generated'],
                'account_name': process_data['account_name'],
                'process_name': process_data['process_name'],
                'process_type': process_type
            })
        if not process_list:
            self.histogram_widget.update_plot([], [])
            return

        df = pd.DataFrame(process_list)
        try:
            df['datetime'] = df['time_generated'].apply(lambda x: datetime.fromtimestamp(x))
            df['hour'] = df['datetime'].dt.floor('h')
        except Exception:
            self.histogram_widget.update_plot([], [])
            return

        if filter_username and filter_username != "All Users":
            df = df[df['account_name'] == filter_username]
        if time_from:
            df = df[df['time_generated'] >= time_from]
        if time_to:
            df = df[df['time_generated'] <= time_to]

        if df.empty:
            self.histogram_widget.update_plot([], [])
            return

        counts = df.groupby('hour').size()
        range_start = pd.Timestamp(datetime.fromtimestamp(time_from)).floor('h')
        range_end = pd.Timestamp(datetime.fromtimestamp(time_to)).ceil('h')
        counts = counts.reindex(
            pd.date_range(
                start=range_start,
                end=range_end,
                freq='h'
            ),
            fill_value=0
        )

        min_date = counts.index.min().date()
        max_date = counts.index.max().date()
        use_date = (max_date - min_date).days > 0
        try:
            if use_date:
                hours = [h.strftime('%m-%d %H:%M') for h in counts.index]
            else:
                hours = [h.strftime('%H:%M') for h in counts.index]
            count_values = counts.values
        except Exception:
            self.histogram_widget.update_plot([], [])
            return

        self.histogram_widget.update_plot(hours, count_values)

    def clear_search_highlights(self):
        self.tree_delegate.set_search_text("")
        self.tree_view.viewport().update()

    def search_and_highlight(self):
        search_text = self.search_input.text().strip()
        self.tree_delegate.set_search_text(search_text)
        self.tree_view.viewport().update()

    def build_process_tree(self, filter_username=None, time_from=None, time_to=None):
        self.tree_model.removeRows(0, self.tree_model.rowCount())
        self.stats_model.removeRows(0, self.stats_model.rowCount())

        children = {}
        for pid, data in self.processes.items():
            creator_pid = data["creator_pid"]
            if creator_pid not in children:
                children[creator_pid] = []
            children[creator_pid].append(pid)

        total_processes = 0
        app_counts = {}
        unique_users = set()
        elevated_processes = 0
        parent_processes = set()
        suspicious_commands = 0
        network_commands = 0
        registry_commands = 0
        filtered_pids = []

        root = QStandardItem("Root")
        for pid, data in self.processes.items():
            if filter_username and data["account_name"] != filter_username:
                continue
            if time_from and data["time_generated"] < time_from:
                continue
            if time_to and data["time_generated"] > time_to:
                continue
            if self.selected_hour:
                process_time = datetime.fromtimestamp(data["time_generated"])
                process_time_floored = process_time.replace(minute=0, second=0, microsecond=0)
                process_hour = process_time_floored.strftime('%H:%M')
                if process_hour != self.selected_hour:
                    continue
            total_processes += 1
            app_name = data["process_name"]
            app_counts[app_name] = app_counts.get(app_name, 0) + 1
            unique_users.add(data["account_name"])
            if data["elevation_type"] in ["TokenElevationTypeFull (3)", "TokenElevationTypeLimited (2)"]:
                elevated_processes += 1
            parent_processes.add(data["creator_name"])
            cmdline = data["cmdline"]
            if any(pattern in cmdline for pattern in self.suspicious_patterns):
                suspicious_commands += 1
            if any(pattern in cmdline for pattern in self.network_patterns):
                network_commands += 1
            if any(pattern in cmdline for pattern in self.registry_patterns):
                registry_commands += 1
            filtered_pids.append(pid)
            if data["creator_pid"] not in self.processes:
                self.add_process_node(root, pid, children)

        stats_rows = []
        total_item = QStandardItem("Total Processes")
        total_count = QStandardItem(str(total_processes))
        total_count.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        stats_rows.append([total_item, total_count])
        if total_processes > 0:
            sorted_apps = sorted(app_counts.items(), key=lambda x: x[0])
            for app, count in sorted_apps:
                app_item = QStandardItem(app)
                count_item = QStandardItem(str(count))
                count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if any(data["full_path"] in self.flagged_paths for data in self.processes.values()
                       if data["process_name"] == app):
                    app_item.setForeground(QBrush(QColor(0, 0, 255)))
                stats_rows.append([app_item, count_item])
        for row in stats_rows:
            self.stats_model.appendRow(row)

        self.resize_stats_table()

        self.unique_users_label.setText(f"Unique Users: {len(unique_users)}")
        self.elevated_processes_label.setText(f"Elevated Processes: {elevated_processes}")
        self.unique_parents_label.setText(f"Unique Parent Processes: {len(parent_processes)}")
        self.suspicious_commands_label.setText(f"Suspicious Command Lines: {suspicious_commands}")
        self.network_commands_label.setText(f"Network-Related Commands: {network_commands}")
        self.registry_commands_label.setText(f"Registry-Related Commands: {registry_commands}")

        if root.hasChildren():
            self.tree_model.appendRow(root)
            self.tree_view.expandAll()

        self.build_process_hist(filter_username, time_from, time_to)

        font_metrics = QFontMetrics(self.tree_view.font())
        max_cmdline_width = self.tree_view.columnWidth(5)
        for row in range(self.tree_model.rowCount()):
            index = self.tree_model.index(row, 5)
            item = self.tree_model.itemFromIndex(index)
            if item:
                text = item.text()
                text_width = font_metrics.horizontalAdvance(text)
                max_cmdline_width = max(max_cmdline_width, text_width + 30)
            for child_row in range(self.tree_model.rowCount(index)):
                child_index = self.tree_model.index(child_row, 5, index)
                child_item = self.tree_model.itemFromIndex(child_index)
                if child_item:
                    text = child_item.text()
                    text_width = font_metrics.horizontalAdvance(text)
                    max_cmdline_width = max(max_cmdline_width, text_width + 30)

        self.tree_view.setColumnWidth(5, max_cmdline_width)

    def add_process_node(self, parent, pid, children):
        data = self.processes[pid]
        node = QStandardItem(data["process_name"])
        full_path_item = QStandardItem(data["full_path"])
        full_path_item.setData(data["full_path"], Qt.ToolTipRole)
        account_item = QStandardItem(data["account_name"])
        pid_item = QStandardItem(pid)
        try:
            time_str = (datetime.fromtimestamp(data["time_generated"]).strftime('%Y-%m-%d %H:%M:%S')
                        if data["time_generated"] > 0 else "Unknown")
        except (OSError, ValueError, TypeError):
            time_str = "Unknown"
        created_time_item = QStandardItem(time_str)
        cmdline_display = data["cmdline"]
        cmdline_item = QStandardItem(cmdline_display)
        cmdline_item.setData(data["cmdline"], Qt.ToolTipRole)

        if data["full_path"] in self.flagged_paths:
            blue_brush = QBrush(QColor(0, 0, 255))
            for item in [node, full_path_item, account_item, pid_item, created_time_item, cmdline_item]:
                item.setForeground(blue_brush)
        elif data["elevation_type"] in ["TokenElevationTypeFull (3)", "TokenElevationTypeLimited (2)"]:
            red_brush = QBrush(QColor(200, 0, 0))
            for item in [node, full_path_item, account_item, pid_item, created_time_item, cmdline_item]:
                item.setForeground(red_brush)

        parent.appendRow([node, full_path_item, account_item, pid_item, created_time_item, cmdline_item])

        for child_pid in children.get(pid, []):
            if (self.processes[child_pid]["account_name"] == self.user_combo.currentText() or
                    self.user_combo.currentText() == "All Users"):
                self.add_process_node(node, child_pid, children)

    def filter_tree(self, username=None):
        if not self.processes:
            return

        start_value = self.start_slider.value()
        end_value = self.end_slider.value()
        start_time = self.min_time + (start_value * self.step_size)
        end_time = self.min_time + (end_value * self.step_size)

        if start_time > end_time:
            self.stats_model.removeRows(0, self.stats_model.rowCount())
            self.stats_model.appendRow([QStandardItem("Invalid time range"), QStandardItem("")])
            self.unique_users_label.setText("Unique Users: 0")
            self.elevated_processes_label.setText("Elevated Processes: 0")
            self.unique_parents_label.setText("Unique Parent Processes: 0")
            self.suspicious_commands_label.setText("Suspicious Command Lines: 0")
            self.network_commands_label.setText("Network-Related Commands: 0")
            self.registry_commands_label.setText("Registry-Related Commands: 0")
            self.histogram_widget.update_plot([], [])
            QMessageBox.warning(
                self, "Invalid Time Range", "The start time cannot be later than the end time."
            )
            return

        username = self.user_combo.currentText() if username is None else username
        filter_username = username if username != "All Users" else None
        self.build_process_tree(filter_username, start_time, end_time)

        self.search_and_highlight()


def main():
    app = QApplication(sys.argv)
    window = ProcessTreeWindow()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()