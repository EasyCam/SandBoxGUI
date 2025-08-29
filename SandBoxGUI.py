#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows Sandbox Configuration Tool

A GUI tool for creating and managing Windows Sandbox configuration files (.wsb)
Based on Microsoft Windows Sandbox documentation.

Author: Generated for SandboxGUI
Date: 2025-08-29
"""

import sys
import os
import json
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox,
    QComboBox, QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QMessageBox, QSpinBox, QFormLayout, QSplitter,
    QScrollArea, QFrame, QGridLayout
)
from PySide6.QtCore import Qt, QSettings, Signal
from PySide6.QtGui import QIcon, QFont, QAction, QPixmap


class MappedFolder:
    """Represents a mapped folder configuration"""
    def __init__(self, host_folder="", sandbox_folder="", read_only=True):
        self.host_folder = host_folder
        self.sandbox_folder = sandbox_folder
        self.read_only = read_only

    def to_dict(self):
        return {
            'host_folder': self.host_folder,
            'sandbox_folder': self.sandbox_folder,
            'read_only': self.read_only
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            data.get('host_folder', ''),
            data.get('sandbox_folder', ''),
            data.get('read_only', True)
        )


class MappedFoldersWidget(QWidget):
    """Widget for managing mapped folders"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.mapped_folders = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Controls
        controls_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Folder")
        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.setEnabled(False)
        
        controls_layout.addWidget(self.add_button)
        controls_layout.addWidget(self.remove_button)
        controls_layout.addStretch()
        
        # Table
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Host Folder", "Sandbox Folder", "Read Only"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        
        layout.addLayout(controls_layout)
        layout.addWidget(self.table)
        
        # Connect signals
        self.add_button.clicked.connect(self.add_folder)
        self.remove_button.clicked.connect(self.remove_folder)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.cellChanged.connect(self.on_cell_changed)

    def add_folder(self):
        """Add a new mapped folder"""
        dialog = QFileDialog()
        folder = dialog.getExistingDirectory(self, "Select Host Folder")
        if folder:
            mapped_folder = MappedFolder(folder, "C:\\Users\\WDAGUtilityAccount\\Desktop\\Shared")
            self.mapped_folders.append(mapped_folder)
            self.refresh_table()

    def remove_folder(self):
        """Remove selected mapped folder"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            del self.mapped_folders[current_row]
            self.refresh_table()

    def on_selection_changed(self):
        """Handle selection change"""
        has_selection = len(self.table.selectedItems()) > 0
        self.remove_button.setEnabled(has_selection)

    def on_cell_changed(self, row, column):
        """Handle cell value changes"""
        if row < len(self.mapped_folders):
            item = self.table.item(row, column)
            if item and column == 0:  # Host folder
                self.mapped_folders[row].host_folder = item.text()
            elif item and column == 1:  # Sandbox folder
                self.mapped_folders[row].sandbox_folder = item.text()

    def refresh_table(self):
        """Refresh the table with current mapped folders"""
        self.table.setRowCount(len(self.mapped_folders))
        
        for row, folder in enumerate(self.mapped_folders):
            # Host folder
            self.table.setItem(row, 0, QTableWidgetItem(folder.host_folder))
            
            # Sandbox folder
            self.table.setItem(row, 1, QTableWidgetItem(folder.sandbox_folder))
            
            # Read only checkbox
            checkbox = QCheckBox()
            checkbox.setChecked(folder.read_only)
            checkbox.stateChanged.connect(lambda state, r=row: self.on_readonly_changed(r, state))
            self.table.setCellWidget(row, 2, checkbox)

    def on_readonly_changed(self, row, state):
        """Handle read-only checkbox change"""
        if row < len(self.mapped_folders):
            self.mapped_folders[row].read_only = state == Qt.CheckState.Checked

    def get_folders(self):
        """Get all mapped folders"""
        return self.mapped_folders

    def set_folders(self, folders):
        """Set mapped folders"""
        self.mapped_folders = folders
        self.refresh_table()


class SandboxConfigTool(QMainWindow):
    """Main application window for Windows Sandbox configuration"""
    
    def __init__(self):
        super().__init__()
        self.settings = QSettings("SandboxGUI", "WindowsSandboxConfig")
        self.current_file = None
        self.setup_ui()
        self.setup_menu()
        self.load_settings()
        
    def setup_ui(self):
        """Setup the user interface"""
        self.setWindowTitle("Windows Sandbox Configuration Tool")
        self.setMinimumSize(800, 600)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create tabs
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        self.setup_general_tab()
        self.setup_folders_tab()
        self.setup_startup_tab()
        self.setup_preview_tab()
        
        # Status bar
        self.statusBar().showMessage("Ready")
        
    def setup_menu(self):
        """Setup the menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        new_action = QAction("New", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_config)
        file_menu.addAction(new_action)
        
        open_action = QAction("Open...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_config)
        file_menu.addAction(open_action)
        
        save_action = QAction("Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_config)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save As...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_config_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("Export WSB...", self)
        export_action.triggered.connect(self.export_wsb)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Help menu
        help_menu = menubar.addMenu("Help")
        
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
    def setup_general_tab(self):
        """Setup the general configuration tab"""
        tab = QWidget()
        self.tab_widget.addTab(tab, "General")
        
        layout = QVBoxLayout(tab)
        
        # VGpu settings
        vgpu_group = QGroupBox("Virtual GPU")
        vgpu_layout = QFormLayout(vgpu_group)
        
        self.vgpu_enabled = QCheckBox("Enable vGPU")
        self.vgpu_enabled.setChecked(True)
        self.vgpu_enabled.setToolTip("Enable or disable virtualized GPU support")
        vgpu_layout.addRow(self.vgpu_enabled)
        
        # Networking settings
        network_group = QGroupBox("Networking")
        network_layout = QFormLayout(network_group)
        
        self.networking_enabled = QCheckBox("Enable Networking")
        self.networking_enabled.setChecked(True)
        self.networking_enabled.setToolTip("Enable or disable network access")
        network_layout.addRow(self.networking_enabled)
        
        # Audio input settings
        audio_group = QGroupBox("Audio")
        audio_layout = QFormLayout(audio_group)
        
        self.audio_input_enabled = QCheckBox("Enable Audio Input")
        self.audio_input_enabled.setChecked(False)
        self.audio_input_enabled.setToolTip("Enable or disable microphone access")
        audio_layout.addRow(self.audio_input_enabled)
        
        # Video input settings
        video_group = QGroupBox("Video")
        video_layout = QFormLayout(video_group)
        
        self.video_input_enabled = QCheckBox("Enable Video Input")
        self.video_input_enabled.setChecked(False)
        self.video_input_enabled.setToolTip("Enable or disable camera access")
        video_layout.addRow(self.video_input_enabled)
        
        # Protected client settings
        protected_group = QGroupBox("Security")
        protected_layout = QFormLayout(protected_group)
        
        self.protected_client_enabled = QCheckBox("Enable Protected Client")
        self.protected_client_enabled.setChecked(False)
        self.protected_client_enabled.setToolTip("Enable additional security protections")
        protected_layout.addRow(self.protected_client_enabled)
        
        # Printer redirection settings
        printer_group = QGroupBox("Printer")
        printer_layout = QFormLayout(printer_group)
        
        self.printer_redirection_enabled = QCheckBox("Enable Printer Redirection")
        self.printer_redirection_enabled.setChecked(False)
        self.printer_redirection_enabled.setToolTip("Enable or disable printer access")
        printer_layout.addRow(self.printer_redirection_enabled)
        
        # Clipboard redirection settings
        clipboard_group = QGroupBox("Clipboard")
        clipboard_layout = QFormLayout(clipboard_group)
        
        self.clipboard_redirection_enabled = QCheckBox("Enable Clipboard Redirection")
        self.clipboard_redirection_enabled.setChecked(True)
        self.clipboard_redirection_enabled.setToolTip("Enable or disable clipboard sharing")
        clipboard_layout.addRow(self.clipboard_redirection_enabled)
        
        # Memory settings
        memory_group = QGroupBox("Memory")
        memory_layout = QFormLayout(memory_group)
        
        self.memory_mb = QSpinBox()
        self.memory_mb.setRange(512, 32768)
        self.memory_mb.setValue(4096)
        self.memory_mb.setSuffix(" MB")
        self.memory_mb.setToolTip("Memory allocation in megabytes")
        memory_layout.addRow("Memory Size:", self.memory_mb)
        
        # Hostname settings
        hostname_group = QGroupBox("Hostname")
        hostname_layout = QFormLayout(hostname_group)
        
        self.hostname_enabled = QCheckBox("Set Custom Hostname")
        self.hostname_enabled.setChecked(False)
        self.hostname_enabled.setToolTip("Enable custom hostname for the sandbox")
        hostname_layout.addRow(self.hostname_enabled)
        
        self.hostname_value = QLineEdit()
        self.hostname_value.setPlaceholderText("e.g., MySandbox")
        self.hostname_value.setEnabled(False)
        self.hostname_value.setToolTip("Custom hostname for the sandbox (max 15 characters)")
        self.hostname_value.setMaxLength(15)
        hostname_layout.addRow("Hostname:", self.hostname_value)
        
        # Connect hostname checkbox to enable/disable input
        self.hostname_enabled.toggled.connect(self.hostname_value.setEnabled)
        
        # Appearance settings
        appearance_group = QGroupBox("Appearance")
        appearance_layout = QFormLayout(appearance_group)
        
        self.force_dark_mode = QCheckBox("Force Dark Mode")
        self.force_dark_mode.setChecked(False)
        self.force_dark_mode.setToolTip("Force the sandbox to use dark theme")
        appearance_layout.addRow(self.force_dark_mode)
        
        # Add all groups to layout
        layout.addWidget(vgpu_group)
        layout.addWidget(network_group)
        layout.addWidget(audio_group)
        layout.addWidget(video_group)
        layout.addWidget(protected_group)
        layout.addWidget(printer_group)
        layout.addWidget(clipboard_group)
        layout.addWidget(memory_group)
        layout.addWidget(hostname_group)
        layout.addWidget(appearance_group)
        layout.addStretch()
        
    def setup_folders_tab(self):
        """Setup the mapped folders tab"""
        tab = QWidget()
        self.tab_widget.addTab(tab, "Mapped Folders")
        
        layout = QVBoxLayout(tab)
        
        # Instructions
        instructions = QLabel(
            "Configure folders to be shared between the host and sandbox.\n"
            "Host folders will be accessible from within the sandbox."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)
        
        # Mapped folders widget
        self.mapped_folders_widget = MappedFoldersWidget()
        layout.addWidget(self.mapped_folders_widget)
        
    def setup_startup_tab(self):
        """Setup the startup command tab"""
        tab = QWidget()
        self.tab_widget.addTab(tab, "Startup")
        
        layout = QVBoxLayout(tab)
        
        # Logon command group
        logon_group = QGroupBox("Logon Command")
        logon_layout = QVBoxLayout(logon_group)
        
        instructions = QLabel(
            "Specify a command to run when the sandbox starts up.\n"
            "This command will be executed after the user logs in."
        )
        instructions.setWordWrap(True)
        logon_layout.addWidget(instructions)
        
        # Command input
        command_layout = QFormLayout()
        
        self.logon_command = QLineEdit()
        self.logon_command.setPlaceholderText("e.g., C:\\Windows\\System32\\cmd.exe")
        command_layout.addRow("Command:", self.logon_command)
        
        logon_layout.addLayout(command_layout)
        
        # Example commands
        examples_group = QGroupBox("Example Commands")
        examples_layout = QVBoxLayout(examples_group)
        
        examples_text = QTextEdit()
        examples_text.setReadOnly(True)
        examples_text.setMaximumHeight(150)
        examples_text.setPlainText(
            "Open Command Prompt:\n"
            "C:\\Windows\\System32\\cmd.exe\n\n"
            "Open PowerShell:\n"
            "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe\n\n"
            "Open specific application:\n"
            "C:\\Users\\WDAGUtilityAccount\\Desktop\\Shared\\myapp.exe\n\n"
            "Run batch file:\n"
            "C:\\Users\\WDAGUtilityAccount\\Desktop\\Shared\\setup.bat"
        )
        examples_layout.addWidget(examples_text)
        
        layout.addWidget(logon_group)
        layout.addWidget(examples_group)
        layout.addStretch()
        
    def setup_preview_tab(self):
        """Setup the preview tab"""
        tab = QWidget()
        self.tab_widget.addTab(tab, "Preview")
        
        layout = QVBoxLayout(tab)
        
        # Preview controls
        controls_layout = QHBoxLayout()
        
        refresh_button = QPushButton("Refresh Preview")
        refresh_button.clicked.connect(self.update_preview)
        controls_layout.addWidget(refresh_button)
        controls_layout.addStretch()
        
        layout.addLayout(controls_layout)
        
        # Preview text
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Courier New", 10))
        layout.addWidget(self.preview_text)
        
        # Auto-update preview when tab is selected
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        
    def on_tab_changed(self, index):
        """Handle tab change"""
        if self.tab_widget.tabText(index) == "Preview":
            self.update_preview()
            
    def generate_wsb_xml(self):
        """Generate WSB XML configuration"""
        root = Element("Configuration")
        
        # VGpu
        if not self.vgpu_enabled.isChecked():
            vgpu = SubElement(root, "VGpu")
            vgpu.text = "Disable"
            
        # Networking
        if not self.networking_enabled.isChecked():
            networking = SubElement(root, "Networking")
            networking.text = "Disable"
            
        # AudioInput
        if self.audio_input_enabled.isChecked():
            audio_input = SubElement(root, "AudioInput")
            audio_input.text = "Enable"
        else:
            audio_input = SubElement(root, "AudioInput")
            audio_input.text = "Disable"
            
        # VideoInput
        if self.video_input_enabled.isChecked():
            video_input = SubElement(root, "VideoInput")
            video_input.text = "Enable"
        else:
            video_input = SubElement(root, "VideoInput")
            video_input.text = "Disable"
            
        # ProtectedClient
        if self.protected_client_enabled.isChecked():
            protected_client = SubElement(root, "ProtectedClient")
            protected_client.text = "Enable"
        else:
            protected_client = SubElement(root, "ProtectedClient")
            protected_client.text = "Disable"
            
        # PrinterRedirection
        if self.printer_redirection_enabled.isChecked():
            printer = SubElement(root, "PrinterRedirection")
            printer.text = "Enable"
        else:
            printer = SubElement(root, "PrinterRedirection")
            printer.text = "Disable"
            
        # ClipboardRedirection
        if not self.clipboard_redirection_enabled.isChecked():
            clipboard = SubElement(root, "ClipboardRedirection")
            clipboard.text = "Disable"
            
        # MemoryInMB
        if self.memory_mb.value() != 4096:  # Only include if not default
            memory = SubElement(root, "MemoryInMB")
            memory.text = str(self.memory_mb.value())
            
        # MappedFolders
        folders = self.mapped_folders_widget.get_folders()
        if folders:
            mapped_folders = SubElement(root, "MappedFolders")
            for folder in folders:
                if folder.host_folder and folder.sandbox_folder:
                    mapped_folder = SubElement(mapped_folders, "MappedFolder")
                    
                    host_folder = SubElement(mapped_folder, "HostFolder")
                    host_folder.text = folder.host_folder
                    
                    sandbox_folder = SubElement(mapped_folder, "SandboxFolder")
                    sandbox_folder.text = folder.sandbox_folder
                    
                    read_only = SubElement(mapped_folder, "ReadOnly")
                    read_only.text = "true" if folder.read_only else "false"
                    
        # LogonCommand
        if self.logon_command.text().strip():
            logon_command = SubElement(root, "LogonCommand")
            command = SubElement(logon_command, "Command")
            command.text = self.logon_command.text().strip()
            
        # HostName
        if self.hostname_enabled.isChecked() and self.hostname_value.text().strip():
            hostname = SubElement(root, "HostName")
            hostname.text = self.hostname_value.text().strip()
            
        # WindowsAppTheme (Dark Mode)
        if self.force_dark_mode.isChecked():
            theme = SubElement(root, "WindowsAppTheme")
            theme.text = "Dark"
            
        return root
        
    def format_xml(self, element):
        """Format XML with proper indentation"""
        rough_string = tostring(element, 'unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")[23:]  # Remove first line
        
    def update_preview(self):
        """Update the preview text"""
        try:
            xml_root = self.generate_wsb_xml()
            formatted_xml = self.format_xml(xml_root)
            self.preview_text.setPlainText(formatted_xml)
        except Exception as e:
            self.preview_text.setPlainText(f"Error generating preview: {str(e)}")
            
    def new_config(self):
        """Create a new configuration"""
        reply = QMessageBox.question(
            self, "New Configuration",
            "Create a new configuration? Any unsaved changes will be lost.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.reset_to_defaults()
            self.current_file = None
            self.update_window_title()
            self.statusBar().showMessage("New configuration created")
            
    def reset_to_defaults(self):
        """Reset all settings to defaults"""
        self.vgpu_enabled.setChecked(True)
        self.networking_enabled.setChecked(True)
        self.audio_input_enabled.setChecked(False)
        self.video_input_enabled.setChecked(False)
        self.protected_client_enabled.setChecked(False)
        self.printer_redirection_enabled.setChecked(False)
        self.clipboard_redirection_enabled.setChecked(True)
        self.memory_mb.setValue(4096)
        self.logon_command.clear()
        self.mapped_folders_widget.set_folders([])
        self.hostname_enabled.setChecked(False)
        self.hostname_value.clear()
        self.force_dark_mode.setChecked(False)
        
    def open_config(self):
        """Open a configuration file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Configuration",
            "", "JSON files (*.json);;All files (*.*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                self.load_configuration(config)
                self.current_file = file_path
                self.update_window_title()
                self.statusBar().showMessage(f"Opened: {file_path}")
                
            except Exception as e:
                QMessageBox.critical(
                    self, "Error",
                    f"Failed to open configuration:\n{str(e)}"
                )
                
    def save_config(self):
        """Save the current configuration"""
        if self.current_file:
            self.save_to_file(self.current_file)
        else:
            self.save_config_as()
            
    def save_config_as(self):
        """Save the configuration to a new file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Configuration",
            "", "JSON files (*.json);;All files (*.*)"
        )
        
        if file_path:
            self.save_to_file(file_path)
            
    def save_to_file(self, file_path):
        """Save configuration to specified file"""
        try:
            config = self.get_current_configuration()
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
                
            self.current_file = file_path
            self.update_window_title()
            self.statusBar().showMessage(f"Saved: {file_path}")
            
        except Exception as e:
            QMessageBox.critical(
                self, "Error",
                f"Failed to save configuration:\n{str(e)}"
            )
            
    def export_wsb(self):
        """Export as WSB file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export WSB Configuration",
            "", "Windows Sandbox files (*.wsb);;All files (*.*)"
        )
        
        if file_path:
            try:
                xml_root = self.generate_wsb_xml()
                formatted_xml = self.format_xml(xml_root)
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(formatted_xml)
                    
                self.statusBar().showMessage(f"Exported WSB: {file_path}")
                
                # Ask if user wants to run the sandbox
                reply = QMessageBox.question(
                    self, "Export Complete",
                    f"WSB file exported successfully!\n\nDo you want to open the sandbox now?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    os.startfile(file_path)
                    
            except Exception as e:
                QMessageBox.critical(
                    self, "Error",
                    f"Failed to export WSB file:\n{str(e)}"
                )
                
    def get_current_configuration(self):
        """Get the current configuration as a dictionary"""
        return {
            'vgpu_enabled': self.vgpu_enabled.isChecked(),
            'networking_enabled': self.networking_enabled.isChecked(),
            'audio_input_enabled': self.audio_input_enabled.isChecked(),
            'video_input_enabled': self.video_input_enabled.isChecked(),
            'protected_client_enabled': self.protected_client_enabled.isChecked(),
            'printer_redirection_enabled': self.printer_redirection_enabled.isChecked(),
            'clipboard_redirection_enabled': self.clipboard_redirection_enabled.isChecked(),
            'memory_mb': self.memory_mb.value(),
            'logon_command': self.logon_command.text(),
            'mapped_folders': [folder.to_dict() for folder in self.mapped_folders_widget.get_folders()],
            'hostname_enabled': self.hostname_enabled.isChecked(),
            'hostname_value': self.hostname_value.text(),
            'force_dark_mode': self.force_dark_mode.isChecked()
        }
        
    def load_configuration(self, config):
        """Load configuration from dictionary"""
        self.vgpu_enabled.setChecked(config.get('vgpu_enabled', True))
        self.networking_enabled.setChecked(config.get('networking_enabled', True))
        self.audio_input_enabled.setChecked(config.get('audio_input_enabled', False))
        self.video_input_enabled.setChecked(config.get('video_input_enabled', False))
        self.protected_client_enabled.setChecked(config.get('protected_client_enabled', False))
        self.printer_redirection_enabled.setChecked(config.get('printer_redirection_enabled', False))
        self.clipboard_redirection_enabled.setChecked(config.get('clipboard_redirection_enabled', True))
        self.memory_mb.setValue(config.get('memory_mb', 4096))
        self.logon_command.setText(config.get('logon_command', ''))
        
        # Load hostname settings
        self.hostname_enabled.setChecked(config.get('hostname_enabled', False))
        self.hostname_value.setText(config.get('hostname_value', ''))
        
        # Load appearance settings
        self.force_dark_mode.setChecked(config.get('force_dark_mode', False))
        
        # Load mapped folders
        folders_data = config.get('mapped_folders', [])
        folders = [MappedFolder.from_dict(data) for data in folders_data]
        self.mapped_folders_widget.set_folders(folders)
        
    def update_window_title(self):
        """Update the window title"""
        if self.current_file:
            filename = os.path.basename(self.current_file)
            self.setWindowTitle(f"Windows Sandbox Configuration Tool - {filename}")
        else:
            self.setWindowTitle("Windows Sandbox Configuration Tool")
            
    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self, "About Windows Sandbox Configuration Tool",
            "<h3>Windows Sandbox Configuration Tool</h3>"
            "<p>A GUI tool for creating and managing Windows Sandbox configuration files (.wsb)</p>"
            "<p>Based on Microsoft Windows Sandbox documentation.</p>"
            "<p><b>Features:</b></p>"
            "<ul>"
            "<li>Configure all Windows Sandbox options</li>"
            "<li>Manage mapped folders</li>"
            "<li>Set startup commands</li>"
            "<li>Preview and export WSB files</li>"
            "<li>Save/load configurations</li>"
            "</ul>"
            "<p><b>Supported Options:</b></p>"
            "<ul>"
            "<li>VGpu (Virtual GPU)</li>"
            "<li>Networking</li>"
            "<li>Audio Input</li>"
            "<li>Video Input</li>"
            "<li>Protected Client</li>"
            "<li>Printer Redirection</li>"
            "<li>Clipboard Redirection</li>"
            "<li>Memory Allocation</li>"
            "<li>Custom Hostname</li>"
            "<li>Force Dark Mode</li>"
            "<li>Mapped Folders</li>"
            "<li>Logon Command</li>"
            "</ul>"
            "<p>Version 1.0</p>"
        )
        
    def load_settings(self):
        """Load application settings"""
        # Restore window geometry
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
            
        # Restore window state
        state = self.settings.value("windowState")
        if state:
            self.restoreState(state)
            
    def save_settings(self):
        """Save application settings"""
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        
    def closeEvent(self, event):
        """Handle application close event"""
        self.save_settings()
        event.accept()


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Windows Sandbox Configuration Tool")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("SandboxGUI")
    app.setOrganizationDomain("sandboxgui.local")
    
    # Create and show main window
    window = SandboxConfigTool()
    window.show()
    
    # Run application
    sys.exit(app.exec())


if __name__ == "__main__":
    main()