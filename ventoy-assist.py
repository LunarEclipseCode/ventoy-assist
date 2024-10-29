import os, ctypes, sys, json, shutil, math, pythoncom, win32com.client, random
from PyQt6 import QtWidgets, QtCore, QtGui
from PyQt6.QtWidgets import QMessageBox, QTableWidgetItem, QCheckBox, QCompleter, QGroupBox
from PyQt6.QtGui import QFont, QPixmap
from PIL import Image
from screeninfo import get_monitors
from pathlib import Path
from skimage.metrics import structural_similarity as ssim
import numpy as np

if getattr(sys, "frozen", False):
    icon_base_path = os.path.dirname(sys.executable)
else:
    icon_base_path = os.path.dirname(os.path.abspath(__file__))

ICON_DIR = os.path.join(icon_base_path, "icons")


# Check if the JSON is valid by trying to load it
def check_json_syntax(file_path):
    try:
        with open(file_path, "r") as file:
            json.load(file)
        return True
    except json.JSONDecodeError:
        return False


# Find image files (iso, wim, img, vhd, vhdx)
def find_image_files(drive_letter, extensions):
    matching_files = []
    for root, dirs, files in os.walk(drive_letter):
        for file in files:
            if file.lower().endswith(extensions):
                matching_files.append(os.path.join(root, file))
    return matching_files


# Get external drives (USB drives and external HDD/SSD)
def get_external_drives():
    pythoncom.CoInitialize()  # Initialize COM threading
    c = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    wmi = c.ConnectServer(".", "root\\cimv2")

    drives = []
    for disk in wmi.ExecQuery("SELECT * FROM Win32_DiskDrive"):
        media_type = disk.MediaType or ""
        model = disk.Model or ""
        interface_type = disk.InterfaceType or ""  # USB, SCSI, IDE, etc.
        disk_size = int(disk.Size) if disk.Size else 0

        # Determine if the drive is external
        is_external = False
        if "USB" in interface_type.upper():
            is_external = True
        elif "EXTERNAL" in media_type.upper():
            is_external = True
        elif "REMOVABLE" in media_type.upper():
            is_external = True

        if is_external:
            partitions = disk.Associators_("Win32_DiskDriveToDiskPartition")
            for partition in partitions:
                logical_disks = partition.Associators_("Win32_LogicalDiskToPartition")
                for logical_disk in logical_disks:
                    drive_letter = logical_disk.DeviceID
                    volume_name = logical_disk.VolumeName  # label
                    logical_size = int(logical_disk.Size) if logical_disk.Size else 0  # Logical disk size

                    # If the logical disk size is zero, fallback to the physical disk size
                    size_to_use = logical_size if logical_size > 0 else disk_size

                    drives.append(
                        {
                            "drive_letter": drive_letter,
                            "volume_name": volume_name,
                            "size": size_to_use,
                            "model": model,
                            "media_type": media_type,
                        }
                    )
    return drives


# Format size with correct units
def format_size(size_bytes):
    if size_bytes == 0:
        return "0B"
    size_name = ("B", "KB", "MB", "GB", "TB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_name[i]}"


# Get absolute path to resource, works for dev and PyInstaller
def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent.resolve()
    return base_path / relative_path


# Main GUI Application
class VentoyApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.iso_aliases = []  # List to store ISOs/directories and their aliases for rename
        self.init_ui()

    def init_ui(self):
        QtWidgets.QApplication.setStyle("windowsvista")
        self.setWindowTitle("Ventoy Assist")

        # Create main layout
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Create tabs
        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setIconSize(QtCore.QSize(24, 24))
        main_layout.addWidget(self.tabs)

        # Apply Icons Tab
        self.apply_icons_tab = QtWidgets.QWidget()
        self.init_apply_icons_ui()
        self.tabs.addTab(self.apply_icons_tab, "Apply Icons")

        # Rename Tab
        self.rename_tab = QtWidgets.QWidget()
        self.init_rename_ui()
        self.tabs.addTab(self.rename_tab, "Rename")

        # Connect the tab change signal to a slot
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.setLayout(main_layout)

        # Set the initial window size based on the initial tab
        self.on_tab_changed(self.tabs.currentIndex())

        # Apply custom styles
        self.apply_styles()

    def apply_styles(self):
        arrow_svg_path = resource_path("resources/arrow.svg").as_posix()
        combo_box_style = f"""
         QComboBox {{
                background-color: #f8f8f8;
                color: #000000;
                border: 1px solid #bdbdbd;
                padding: 0.25em;
                border-radius: 0.25em;
                padding-left: 0.3em;
            }}

            QComboBox:hover {{
                background-color: #e3e3e3;
            }}
            
            QComboBox::down-arrow {{
                image: url("{arrow_svg_path}");
                width: 0.75em;
                height: 0.75em;
            }}
            
            /* Dropdown List */
            QComboBox QAbstractItemView {{
                color: #000000;
                selection-background-color: #d6d6d6; 
                selection-color: #000000;  
                outline: none;  /* Remove the box around the items */
                border: none;
                border-radius: 0.2em;

            }}

            /* Individual Items */
            QComboBox QAbstractItemView::item {{ 
                color: #000000;    
                background-color: #f8f8f8;  /* Text moves on hover if bg not applied */
                padding: 0.25em;
            }}
            
            QComboBox QAbstractItemView::item:hover {{
                background-color: #e3e3e3;  
            }}

            QComboBox QAbstractItemView::item:selected {{
                background-color: #d6d6d6;
            }}

            QComboBox QAbstractItemView::item:focus {{
                background-color: #d6d6d6;  
            }}
            
            QComboBox::drop-down {{
                width: 25px;
                background: transparent;
                margin-right: 0.2em;
            }}
        """
        self.usb_dropdown.setStyleSheet(combo_box_style)
        self.theme_dropdown.setStyleSheet(combo_box_style)
        self.rename_usb_dropdown.setStyleSheet(combo_box_style)
        self.iso_dropdown.setStyleSheet(combo_box_style)

        line_edit_style = """
            QLineEdit {
                border: 1px solid #ccc;
                padding: 0.25em;
                border-radius: 0.25em;
            }
        """

        self.search_bar.setStyleSheet(line_edit_style)
        self.alias_input.setStyleSheet(line_edit_style)

        table_style = """
            QTableWidget {
                background-color: #f0f0f0;
                border: none;
                border-radius: 0.25em;
                gridline-color: #ccc;
                selection-background-color: #d6d6d6;
                selection-color: #000000;
                alternate-background-color: #f0f0f0;
                font-size: 14px;
            }

            QHeaderView::section {
                background-color: #e0e0e0;
                border: none;
                padding: 0.25em;
                font-weight: bold;
                color: #000000;
            }

            QTableWidget {
                outline: none;
            }

            QTableWidget::item {
                padding: 0.25em;
                background-color: #f0f0f0;  
                color: #000000;  
            }

            QTableWidget::item:hover {
                background-color: #e3e3e3;
            }

            QTableWidget::item:selected {
                background-color: #d6d6d6;
                color: #000000;
            }

            QTableWidget::item:focus {
                background-color: #d6d6d6;
                color: #000000;
            }

            /* Alternating Row Colors */
            QTableWidget::item:alternate {
                background-color: #ff0000;
            }
        """
        self.rename_table.setStyleSheet(table_style)

    def on_tab_changed(self, index):
        if index == self.tabs.indexOf(self.apply_icons_tab):
            self.setFixedSize(610, 550)
        elif index == self.tabs.indexOf(self.rename_tab):
            self.setFixedSize(610, 600)

    # Initialize the Apply Icons tab UI
    def init_apply_icons_ui(self):
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(15)

        # GroupBox for USB Drive and Theme selection
        usb_theme_group = QGroupBox("Select USB Drive and Theme")
        usb_theme_layout = QtWidgets.QGridLayout()
        usb_theme_layout.setSpacing(10)

        # USB Drive dropdown
        self.usb_label = QtWidgets.QLabel("USB Drive:")
        usb_theme_layout.addWidget(self.usb_label, 0, 0)

        self.usb_dropdown = QtWidgets.QComboBox()
        self.usb_dropdown.setMinimumWidth(200)
        self.usb_dropdown.setToolTip("Select the USB drive where Ventoy is installed")
        usb_theme_layout.addWidget(self.usb_dropdown, 0, 1)

        # Theme dropdown
        self.theme_label = QtWidgets.QLabel("Theme:")
        usb_theme_layout.addWidget(self.theme_label, 1, 0)

        self.theme_dropdown = QtWidgets.QComboBox()
        self.theme_dropdown.setMinimumWidth(200)
        self.theme_dropdown.setToolTip("Select the Ventoy theme to apply icons to")
        usb_theme_layout.addWidget(self.theme_dropdown, 1, 1)

        usb_theme_group.setLayout(usb_theme_layout)
        layout.addWidget(usb_theme_group)

        # Options GroupBox
        options_group = QGroupBox("Options")
        options_layout = QtWidgets.QVBoxLayout()
        options_layout.setSpacing(5)

        self.apply_all_themes_checkbox = QCheckBox("Apply icons to all themes")
        options_layout.addWidget(self.apply_all_themes_checkbox)

        self.apply_all_resolutions_checkbox = QCheckBox("Apply icons to all resolutions of the selected theme")
        options_layout.addWidget(self.apply_all_resolutions_checkbox)

        self.use_theme_icons_checkbox = QCheckBox("Use theme's icons folder instead of the default icons")
        options_layout.addWidget(self.use_theme_icons_checkbox)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        self.usb_dropdown.currentIndexChanged.connect(self.auto_load_themes)

        self.populate_usb_dropdown(self.usb_dropdown)

        # Start Process button
        self.start_button = QtWidgets.QPushButton("Start Process")
        self.start_button.clicked.connect(self.start_apply_icons)
        self.start_button.setStyleSheet(
            """
            QPushButton {
                background-color: #0078d7;
                color: white;
                font-weight: bold;
                padding: 6px 12px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """
        )
        layout.addWidget(self.start_button)

        # Text below Start Process button
        self.info_label = QtWidgets.QLabel()
        self.info_label.setTextFormat(QtCore.Qt.TextFormat.RichText)
        self.info_label.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextBrowserInteraction)
        self.info_label.setOpenExternalLinks(True)
        self.info_label.setStyleSheet("font-size: 13px; color: #111;")
        self.info_label.setWordWrap(True)
        self.info_label.setText(self.get_info_text())
        layout.addWidget(self.info_label)

        self.apply_icons_tab.setLayout(layout)

    def get_info_text(self):
        return (
            "If you are encountering '<b>ventoy.json not found</b>' or '<b>No themes found</b>' error, "
            "please ensure that you have applied a theme to your Ventoy USB drive. "
            "You can download themes from <a href='https://www.gnome-look.org/browse?cat=109&ord=latest' style='color: #0101ff'>Gnome-Look</a> "
            "and apply them using the <a href='https://www.ventoy.net/en/plugin_theme.html' style='color: #0101ff'>theme plugin</a> in Ventoy Plugson.<br><br>"
            "The following themes are compatible with this tool, as the included icon set matches their color schemes:"
            "<table style='width: 100%; border-collapse: collapse;'>"
            "<tr>"
            "    <td style='padding-right: 1em; padding-top: 0.5em; padding-bottom: 0.5em;'><a href='https://www.gnome-look.org/p/1569525' style='color: #0101ff'>DedSec</a></td>"
            "    <td style='padding-left: 1em; padding-right: 1em;padding-top: 0.5em; padding-bottom: 0.5em; '><a href='https://www.gnome-look.org/p/1603282' style='color: #0101ff'>Dark Matter</a></td>"
            "    <td style='padding-left: 1em; padding-right: 1em;padding-top: 0.5em; padding-bottom: 0.5em;'><a href='https://www.gnome-look.org/p/1307852' style='color: #0101ff'>TELA</a></td>"
            "    <td style='padding-left: 1em; padding-right: 1em;padding-top: 0.5em; padding-bottom: 0.5em;'><a href='https://www.gnome-look.org/p/1195799' style='color: #0101ff'>Plasma-dark</a></td>"
            "    <td style='padding-left: 1em; padding-right: 1em;padding-top: 0.5em; padding-bottom: 0.5em;'><a href='https://www.gnome-look.org/p/1850334' style='color: #0101ff'>Fate Series</a></td>"
            "</tr>"
            "</table>"
            "If you prefer to use the icon set provided by the theme instead of the one included with this program, simply enable the '<b>Use theme's icons folder</b>' option."
        )

    # Initialize the Rename tab UI
    def init_rename_ui(self):
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(15)

        # GroupBox for USB Drive selection
        usb_group = QGroupBox("Select USB Drive")
        usb_layout = QtWidgets.QHBoxLayout()

        self.rename_usb_label = QtWidgets.QLabel("USB Drive:")
        usb_layout.addWidget(self.rename_usb_label)

        self.rename_usb_dropdown = QtWidgets.QComboBox()
        self.rename_usb_dropdown.setMinimumWidth(200)
        self.rename_usb_dropdown.setToolTip("Select the USB drive where Ventoy is installed")
        usb_layout.addWidget(self.rename_usb_dropdown)

        usb_group.setLayout(usb_layout)
        layout.addWidget(usb_group)

        # GroupBox for Path Selection
        path_group = QGroupBox("Select Path")
        path_layout = QtWidgets.QGridLayout()
        path_layout.setSpacing(10)

        # Search bar with auto-complete
        self.search_label = QtWidgets.QLabel("Type Path:")
        path_layout.addWidget(self.search_label, 0, 0)

        self.search_bar = QtWidgets.QLineEdit()
        self.search_bar.setPlaceholderText("Start typing to search...")
        self.search_bar.setToolTip("Type the path to the ISO or folder")
        path_layout.addWidget(self.search_bar, 0, 1)

        # ISO/Folder Files list (Dropdown)
        self.iso_label = QtWidgets.QLabel("Select from List:")
        path_layout.addWidget(self.iso_label, 1, 0)

        self.iso_dropdown = QtWidgets.QComboBox()
        self.iso_dropdown.setMinimumWidth(200)
        self.iso_dropdown.setToolTip("Select the ISO or folder from the list")
        path_layout.addWidget(self.iso_dropdown, 1, 1)

        path_group.setLayout(path_layout)
        layout.addWidget(path_group)

        # Track the last changed field
        self.last_changed_field = "dropdown"

        # Connect signals
        self.iso_dropdown.currentIndexChanged.connect(self.on_dropdown_changed)
        self.search_bar.textEdited.connect(self.on_search_bar_changed)
        self.rename_usb_dropdown.currentIndexChanged.connect(self.auto_load_paths)

        # GroupBox for alias input
        alias_group = QGroupBox("Set Alias")
        alias_layout = QtWidgets.QHBoxLayout()

        self.alias_label = QtWidgets.QLabel("New Alias:")
        alias_layout.addWidget(self.alias_label)

        self.alias_input = QtWidgets.QLineEdit()
        self.alias_input.setPlaceholderText("Enter new alias")
        self.alias_input.setToolTip("Enter the new alias for the selected path")
        alias_layout.addWidget(self.alias_input)

        alias_group.setLayout(alias_layout)
        layout.addWidget(alias_group)

        # Add to Rename list button
        self.add_button = QtWidgets.QPushButton("Add to Rename List")
        self.add_button.clicked.connect(self.add_to_rename_list)
        self.add_button.setStyleSheet(
            """
            QPushButton {
                background-color: #ffffff;
                color: black;
                font-weight: semi-bold;
                padding: 6px 12px;
                border-radius: 4px;
                border: 1px solid #d3d3d3; 
            }
            QPushButton:hover {
                background-color: #e3e3e3;
            }
        """
        )
        layout.addWidget(self.add_button)

        # Table to display the ISO/dirs and their new aliases
        self.rename_table = QtWidgets.QTableWidget()
        self.rename_table.setColumnCount(2)
        self.rename_table.setHorizontalHeaderLabels(["Path", "New Alias"])
        self.rename_table.verticalHeader().setVisible(False)
        self.rename_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.rename_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked)
        self.rename_table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # Adjust column resize modes
        header = self.rename_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Interactive)

        layout.addWidget(self.rename_table)
        self.rename_table.cellChanged.connect(self.on_alias_cell_changed)

        # Start Rename button
        self.rename_button = QtWidgets.QPushButton("Apply Rename")
        self.rename_button.clicked.connect(self.start_rename)
        self.rename_button.setStyleSheet(
            """
            QPushButton {
                background-color: #0078d7;
                color: white;
                font-weight: bold;
                padding: 6px 12px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """
        )
        layout.addWidget(self.rename_button)

        self.rename_tab.setLayout(layout)

        # Populate the USB dropdown
        self.populate_usb_dropdown(self.rename_usb_dropdown)

    # Load themes from ventoy.json
    def auto_load_themes(self):
        current_index = self.usb_dropdown.currentIndex()
        drive_letter = self.usb_dropdown.itemData(current_index)
        if not drive_letter or "No external drives found" in drive_letter:
            self.theme_dropdown.clear()
            self.theme_dropdown.addItem("No external drives found")
            return

        ventoy_dir = os.path.join(drive_letter + "\\", "ventoy")
        try:
            ventoy_json = self.read_ventoy_json(ventoy_dir)
        except FileNotFoundError:
            self.theme_dropdown.clear()
            self.theme_dropdown.addItem("ventoy.json not found")
            return
        except ValueError:
            self.theme_dropdown.clear()
            self.theme_dropdown.addItem("Invalid ventoy.json syntax")
            return

        # Find keys starting with 'theme' (including 'theme' itself)
        theme_files = set()
        theme_keys = [key for key in ventoy_json.keys() if key.startswith("theme")]

        # For each key, get the 'file' attribute
        for theme_key in theme_keys:
            theme_entry = ventoy_json[theme_key]
            if "file" in theme_entry:
                file_field = theme_entry["file"]
                if isinstance(file_field, str):
                    file_field = [file_field]
                for file_path in file_field:
                    # Construct the full path
                    if file_path.startswith("/"):
                        file_path = file_path[1:]
                    file_path = file_path.replace("/", os.sep)
                    full_path = os.path.join(drive_letter + "\\", file_path)
                    if os.path.exists(full_path):
                        # Extract theme name from the path
                        theme_name = os.path.basename(os.path.dirname(full_path))
                        theme_files.add(theme_name)

        # Update the theme dropdown
        self.theme_dropdown.clear()
        if theme_files:
            for theme_name in sorted(theme_files):
                self.theme_dropdown.addItem(theme_name)
        else:
            self.theme_dropdown.addItem("No themes found")

    # Populate USB drive dropdown
    def populate_usb_dropdown(self, dropdown):
        dropdown.clear()
        usb_drives = get_external_drives()
        if usb_drives:
            for drive in usb_drives:
                drive_letter = drive["drive_letter"]
                volume_name = drive["volume_name"] or "Unknown"
                size = drive["size"]
                model = drive["model"] or ""
                size_str = format_size(size)
                display_text = f"{drive_letter} [{size_str}] {model.strip()}"
                dropdown.addItem(display_text, drive_letter)
        else:
            dropdown.addItem("No external drives found")

    def start_apply_icons(self):
        # Get the selected USB drive
        current_index = self.usb_dropdown.currentIndex()
        drive_letter = self.usb_dropdown.itemData(current_index)
        if not drive_letter or "No external drives found" in drive_letter:
            QMessageBox.critical(self, "Error", "No external drives detected.")
            return

        # Get the selected theme
        selected_theme = self.theme_dropdown.currentText()
        if not selected_theme or selected_theme == "No themes found":
            QMessageBox.critical(self, "Error", "No theme selected.")
            return

        # Determine which themes to apply icons to
        apply_to_all_themes = self.apply_all_themes_checkbox.isChecked()
        apply_to_all_resolutions = self.apply_all_resolutions_checkbox.isChecked()
        use_theme_icons = self.use_theme_icons_checkbox.isChecked()

        ventoy_dir = os.path.join(drive_letter + "\\", "ventoy")
        try:
            ventoy_json = self.read_ventoy_json(ventoy_dir)
        except FileNotFoundError as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        except ValueError as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        theme_paths = self.collect_theme_paths(
            drive_letter,
            ventoy_json,
            selected_theme,
            apply_to_all_themes,
            apply_to_all_resolutions,
        )
        if theme_paths is None:
            QMessageBox.critical(self, "Error", "No matching themes found to apply icons.")
            return

        # Get filenames with specified extensions from the selected drive
        extensions = (".iso", ".wim", ".img", ".vhd", ".vhdx")
        files = find_image_files(drive_letter + "\\", extensions)

        matching_tools = []
        for theme_folder in theme_paths:
            icons_path = os.path.join(theme_folder, "icons")

            if not os.path.exists(icons_path):
                QMessageBox.warning(
                    self,
                    "Warning",
                    f"No icons folder found in theme {os.path.basename(theme_folder)}. Skipping.",
                )
                continue

            if use_theme_icons:
                icon_map = {}
                for icon_file in os.listdir(icons_path):
                    if icon_file.lower().endswith(".png"):
                        icon_name = os.path.splitext(icon_file)[0]
                        icon_map[icon_name] = icon_name
            else:
                png_files = [f for f in os.listdir(icons_path) if f.lower().endswith(".png")]
                icon_size_value = None

                if "ubuntu.png" in png_files:
                    # Use the resolution of ubuntu.png
                    ubuntu_path = os.path.join(icons_path, "ubuntu.png")
                    pixmap = QPixmap(ubuntu_path)
                    if pixmap.isNull():
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"Failed to load ubuntu.png in theme {os.path.basename(theme_folder)}. Using default icon size.",
                        )
                        icon_size_value = self.icon_size_from_res()
                    else:
                        icon_size_value = pixmap.width()  # Assuming square icons
                else:
                    if png_files:
                        selected_file = random.choice(png_files)
                        selected_path = os.path.join(icons_path, selected_file)
                        pixmap = QPixmap(selected_path)
                        if pixmap.isNull():
                            QMessageBox.warning(
                                self,
                                "Warning",
                                f"Failed to load {selected_file} in theme {os.path.basename(theme_folder)}. Using default icon size.",
                            )
                            icon_size_value = self.icon_size_from_res()
                        else:
                            icon_size_value = pixmap.width()
                    else:
                        icon_size_value = self.icon_size_from_res()

                if icon_size_value is None:
                    icon_size_value = self.icon_size_from_res()

                icon_size = (icon_size_value, icon_size_value)

                success, icon_map = self.copy_and_resize_icons(ICON_DIR, icons_path, icon_size)
                if not success:
                    return

            matching_tools.extend(self.get_matching_tools(files, icon_map))

        # Remove duplicates from matching_tools
        matching_tools = list(set(matching_tools))

        # Sort matching_tools by length of key in descending order and then case-insensitive
        matching_tools.sort(key=lambda x: (-len(x[0]), x[0].lower()))

        # Add menu_class entries
        if "menu_class" not in ventoy_json:
            ventoy_json["menu_class"] = []

        for key_string, class_string in matching_tools:
            menu_entry = {"key": key_string, "class": class_string}
            ventoy_json["menu_class"].append(menu_entry)

        # Remove duplicates and sort the menu_class entries
        unique_menu_class = {}
        for entry in ventoy_json["menu_class"]:
            if "key" in entry:
                unique_key = f"key:{entry['key']}"
                sort_key = entry["key"]
            elif "dir" in entry:
                unique_key = f"dir:{entry['dir']}"
                sort_key = entry["dir"]
            else:
                continue
            unique_menu_class[unique_key] = (sort_key, entry)  # Store sort_key for sorting

        sorted_menu_class_entries = [entry for _, entry in sorted(unique_menu_class.values(), key=lambda x: (-len(x[0]), x[0].lower()))]

        # Separate entries where key.lower() == "linux" and move them to the end
        linux_entries = [entry for entry in sorted_menu_class_entries if entry.get("key", "").lower() == "linux"]
        non_linux_entries = [entry for entry in sorted_menu_class_entries if entry.get("key", "").lower() != "linux"]

        # Combine non-linux entries with linux entries
        ventoy_json["menu_class"] = non_linux_entries + linux_entries

        # Save the updated ventoy.json
        self.save_ventoy_json(ventoy_dir, ventoy_json)

    def icon_size_from_res(self):
        try:
            monitors = get_monitors()
            if monitors:
                width = monitors[0].width
                height = monitors[0].height

                # Get the scaling factor
                user32 = ctypes.windll.user32
                hdc = user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)

                # Default DPI is 96
                scale_factor = dpi / 96.0
                scaled_width = int(width / scale_factor)
                scaled_height = int(height / scale_factor)
                resolution = (scaled_width, scaled_height)
            else:
                resolution = (1920, 1080)
        except Exception:
            resolution = (1920, 1080)

        # Mapping of resolutions to icon sizes
        resolution_icon_size_map = {
            (1920, 1080): 32,
            (2560, 1080): 32,
            (2560, 1440): 48,
            (3440, 1440): 48,
            (3840, 2160): 64,
        }

        # Find the closest matching resolution
        min_diff = float("inf")
        icon_size = 32  # Default icon size
        for res, size in resolution_icon_size_map.items():
            diff = abs(resolution[0] - res[0]) + abs(resolution[1] - res[1])
            if diff < min_diff:
                min_diff = diff
                icon_size = size

        return icon_size

    def read_ventoy_json(self, ventoy_dir, is_rename=False):
        ventoy_json_path = os.path.join(ventoy_dir, "ventoy.json")
        if not os.path.exists(ventoy_json_path):
            raise FileNotFoundError("ventoy.json not found")

        try:
            if is_rename:
                with open(ventoy_json_path, "r") as json_file:
                    content = json_file.read().strip()
                    if not content:
                        ventoy_json = {}
                    else:
                        ventoy_json = json.loads(content)
                return ventoy_json
            else:
                with open(ventoy_json_path, "r") as json_file:
                    ventoy_json = json.load(json_file)
                return ventoy_json
        except json.JSONDecodeError:
            raise ValueError("Invalid ventoy.json syntax")

    def save_ventoy_json(self, ventoy_dir, ventoy_json):
        ventoy_json_path = os.path.join(ventoy_dir, "ventoy.json")
        temp_ventoy_json_path = os.path.join(ventoy_dir, "temp_ventoy.json")
        with open(temp_ventoy_json_path, "w") as json_file:
            json.dump(ventoy_json, json_file, indent=4)

        if check_json_syntax(temp_ventoy_json_path):
            shutil.move(temp_ventoy_json_path, ventoy_json_path)
            QMessageBox.information(self, "Success", f"Updated ventoy.json saved at {ventoy_json_path}")
        else:
            os.remove(temp_ventoy_json_path)
            QMessageBox.critical(
                self,
                "Error",
                "Syntax error in the modified ventoy.json. No changes were made.",
            )

    def collect_theme_paths(self, drive_letter, ventoy_json, selected_theme, apply_to_all_themes, apply_to_all_resolutions):
        theme_paths = []

        # Find theme entries in ventoy.json
        theme_keys = [key for key in ventoy_json.keys() if key.startswith("theme")]

        for theme_key in theme_keys:
            theme_entry = ventoy_json[theme_key]
            if "file" in theme_entry:
                file_field = theme_entry["file"]
                if isinstance(file_field, str):
                    file_field = [file_field]
                for file_path in file_field:
                    # Extract theme name from the path
                    if file_path.startswith("/"):
                        file_path = file_path[1:]
                    file_path = file_path.replace("/", os.sep)
                    full_path = os.path.join(drive_letter + "\\", file_path)
                    if os.path.exists(full_path):
                        theme_name = os.path.basename(os.path.dirname(full_path))

                        # Decide whether to include this theme
                        include_theme = False
                        if apply_to_all_themes:
                            include_theme = True
                        elif theme_name == selected_theme:
                            include_theme = True

                        if include_theme:
                            if apply_to_all_resolutions:
                                # Include all resolutions of the theme
                                theme_base_name = theme_name.split("_")[0]
                                theme_dir = os.path.dirname(os.path.dirname(full_path))
                                # Search for all folders starting with the base theme name
                                for dir_name in os.listdir(theme_dir):
                                    if dir_name.startswith(theme_base_name):
                                        theme_paths.append(os.path.join(theme_dir, dir_name))
                            else:
                                theme_paths.append(os.path.dirname(full_path))

        # Remove duplicates
        theme_paths = list(set(theme_paths))

        if not theme_paths:
            return None

        return theme_paths

    def copy_and_resize_icons(self, source_dir, dest_dir, icon_size):
        print(source_dir)
        if not os.path.exists(source_dir):
            QMessageBox.critical(self, "Error", "Local icons folder not found.")
            return False, {}

        icon_map = {}

        for icon_file in os.listdir(source_dir):
            if icon_file.lower().endswith(".png"):
                src_icon_path = os.path.join(source_dir, icon_file)
                icon_name = os.path.splitext(icon_file)[0]  # Original icon name

                # Destination icon path
                dest_icon_filename = icon_file
                dest_icon_path = os.path.join(dest_dir, dest_icon_filename)

                # If a file with the same name exists in the icons folder, rename the dest file to 'name-alt.png'
                if os.path.exists(dest_icon_path):
                    alt_icon_filename = f"{icon_name}-alt.png"
                    alt_icon_path = os.path.join(dest_dir, alt_icon_filename)

                    # Resize and save the alt icon
                    try:
                        with Image.open(src_icon_path) as img:
                            # Convert image to RGBA to ensure transparency is preserved
                            img = img.convert("RGBA")
                            resized_img = img.resize(icon_size, Image.LANCZOS)
                            resized_img.save(alt_icon_path, format="PNG")
                    except Exception as e:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"Failed to process {icon_file}: {e}. Skipping this icon.",
                        )
                        continue

                    # Calculate similarity between original and alternative icons
                    try:
                        with Image.open(dest_icon_path) as original_img, Image.open(alt_icon_path) as alt_img:
                            # Make sure both images have same size
                            original_img = original_img.convert("RGBA").resize(icon_size, Image.LANCZOS)
                            alt_img = alt_img.convert("RGBA").resize(icon_size, Image.LANCZOS)

                            # Convert images to grayscale for SSIM
                            original_gray = original_img.convert("L")
                            alt_gray = alt_img.convert("L")

                            original_array = np.array(original_gray)
                            alt_array = np.array(alt_gray)

                            # Compute SSIM
                            similarity, _ = ssim(original_array, alt_array, full=True)
                            similarity_percentage = similarity * 100

                    except Exception as e:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"Failed to calculate similarity for {icon_name}: {e}. Keeping the '-alt' icon.",
                        )
                        icon_map[icon_name] = os.path.splitext(alt_icon_filename)[0]
                        continue

                    # If similarity is high; remove the '-alt' icon and map to original
                    if similarity_percentage >= 90:
                        try:
                            os.remove(alt_icon_path)
                        except Exception as e:
                            QMessageBox.warning(
                                self,
                                "Warning",
                                f"Failed to remove {alt_icon_filename}: {e}. Keeping the '-alt' icon.",
                            )
                            icon_map[icon_name] = os.path.splitext(alt_icon_filename)[0]
                            continue
                        icon_map[icon_name] = icon_name
                    else:
                        # Similarity is low; keep the '-alt' icon
                        icon_map[icon_name] = os.path.splitext(alt_icon_filename)[0]
                else:
                    # No conflict; copy and resize the icon normally
                    try:
                        with Image.open(src_icon_path) as img:
                            # Convert image to RGBA to ensure transparency is preserved
                            img = img.convert("RGBA")
                            resized_img = img.resize(icon_size, Image.LANCZOS)
                            resized_img.save(dest_icon_path, format="PNG")
                    except Exception as e:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"Failed to process {icon_file}: {e}. Skipping this icon.",
                        )
                        continue

                    icon_map[icon_name] = icon_name

        return True, icon_map

    def get_matching_tools(self, files, icon_map):
        tool_icons = list(icon_map.keys())

        matching_tools = []
        for file in files:
            filename = os.path.basename(file)
            filename_lower = filename.lower()
            matches_for_file = []
            for tool in tool_icons:
                tool_lower = tool.lower()
                index = filename_lower.find(tool_lower)
                if index != -1:
                    # Extract the matching substring from the filename, preserving case
                    matched_string = filename[index : index + len(tool)]
                    matching_tools.append((matched_string, icon_map[tool]))
        return matching_tools

    # Automatically load paths (files and folders) when a USB drive is selected
    def auto_load_paths(self):
        current_index = self.rename_usb_dropdown.currentIndex()
        drive_letter = self.rename_usb_dropdown.itemData(current_index)
        if not drive_letter or "No external drives found" in drive_letter:
            self.iso_dropdown.clear()
            self.iso_dropdown.addItem("No external drives found")
            return

        # Find image files and their parent directories
        extensions = (".iso", ".wim", ".img", ".vhd", ".vhdx")
        image_files = find_image_files(drive_letter + "\\", extensions)

        # Collect unique paths (files and directories)
        paths_set = set()

        for file_path in image_files:
            if "$RECYCLE.BIN" in file_path:
                continue

            paths_set.add(file_path)

            # Add parent directories up to the drive root
            parent_path = os.path.dirname(file_path)
            while parent_path.startswith(drive_letter + "\\"):
                paths_set.add(parent_path)
                parent_path = os.path.dirname(parent_path)
                if parent_path == drive_letter + "\\":
                    break

        # Convert paths to relative for display
        relative_paths = []
        for path in paths_set:
            relative_path = os.path.relpath(path, drive_letter + "\\")
            display_name = relative_path.replace("\\", "/")
            relative_paths.append(display_name)

        relative_paths.sort()

        self.iso_dropdown.blockSignals(True)
        self.iso_dropdown.clear()
        for display_name in relative_paths:
            self.iso_dropdown.addItem(display_name)
        self.iso_dropdown.blockSignals(False)

        # Set up auto-complete for the search bar
        self.completer_model = QtCore.QStringListModel(relative_paths)
        completer = QCompleter(self.completer_model, self.search_bar)
        completer.setCaseSensitivity(QtCore.Qt.CaseSensitivity.CaseInsensitive)
        self.search_bar.setCompleter(completer)

    # Handle dropdown selection change
    def on_dropdown_changed(self):
        if self.last_changed_field == "search_bar":
            return
        self.last_changed_field = "dropdown"
        selected_path = self.iso_dropdown.currentText()
        self.search_bar.blockSignals(True)
        self.search_bar.setText(selected_path)
        self.search_bar.blockSignals(False)

    # Handle search bar text change
    def on_search_bar_changed(self):
        self.last_changed_field = "search_bar"
        # Do not update the dropdown unless the text matches an item
        text = self.search_bar.text()
        index = self.iso_dropdown.findText(text, QtCore.Qt.MatchFlag.MatchExactly)
        if index != -1:
            self.iso_dropdown.blockSignals(True)
            self.iso_dropdown.setCurrentIndex(index)
            self.iso_dropdown.blockSignals(False)

    # Add selected path and alias to the rename list
    def add_to_rename_list(self):
        selected_path = self.search_bar.text() if self.last_changed_field == "search_bar" else self.iso_dropdown.currentText()

        if not selected_path or selected_path == "No external drives found":
            QMessageBox.critical(self, "Error", "No path selected.")
            return

        current_index = self.rename_usb_dropdown.currentIndex()
        drive_letter = self.rename_usb_dropdown.itemData(current_index)
        full_path = os.path.join(drive_letter + "\\", selected_path.replace("/", "\\"))

        if not os.path.exists(full_path):
            QMessageBox.critical(self, "Error", f"Path does not exist: {selected_path}")
            return

        new_alias = self.alias_input.text().strip()

        if not new_alias:
            QMessageBox.critical(self, "Error", "Please enter a new alias.")
            return

        # If the path already exists in the list, remove the old entry
        for index, (path, _) in enumerate(self.iso_aliases):
            if path == selected_path:
                del self.iso_aliases[index]
                break

        # Add the path and alias to the list and update the table
        self.iso_aliases.append((selected_path, new_alias))
        self.update_rename_table()

        # Clear the alias input for the next entry
        self.alias_input.clear()

    # Update the table to display the paths and their new aliases
    def update_rename_table(self):
        self.rename_table.blockSignals(True)  # Prevent recursive signals
        self.rename_table.setRowCount(len(self.iso_aliases))
        for row, (path, alias) in enumerate(self.iso_aliases):
            path_item = QTableWidgetItem(path)
            path_item.setFlags(path_item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            self.rename_table.setItem(row, 0, path_item)

            alias_item = QTableWidgetItem(alias)
            self.rename_table.setItem(row, 1, alias_item)
        self.rename_table.blockSignals(False)  # Re-enable signals

        # Adjust column widths
        total_width = self.rename_table.viewport().width()
        self.rename_table.setColumnWidth(0, int(total_width * 0.7))
        self.rename_table.setColumnWidth(1, int(total_width * 0.3))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "rename_table"):
            total_width = self.rename_table.viewport().width()
            self.rename_table.setColumnWidth(0, int(total_width * 0.7))
            self.rename_table.setColumnWidth(1, int(total_width * 0.3))

    # Handle alias cell changes
    def on_alias_cell_changed(self, row, column):
        if column == 1:
            # Alias column edited
            new_alias = self.rename_table.item(row, column).text()
            path = self.rename_table.item(row, 0).text()
            # Update self.iso_aliases
            for index, (path_item, _) in enumerate(self.iso_aliases):
                if path_item == path:
                    self.iso_aliases[index] = (path_item, new_alias)
                    break

    # Start process for rename
    def start_rename(self):
        current_index = self.rename_usb_dropdown.currentIndex()
        drive_letter = self.rename_usb_dropdown.itemData(current_index)
        if not drive_letter or "No external drives found" in drive_letter:
            QMessageBox.critical(self, "Error", "No external drives detected.")
            return

        ventoy_dir = os.path.join(drive_letter + "\\", "ventoy")
        if not os.path.exists(ventoy_dir):
            os.makedirs(ventoy_dir)

        # Read ventoy.json
        try:
            ventoy_json = self.read_ventoy_json(ventoy_dir, True)
        except FileNotFoundError as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        except ValueError as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        # Add the new alias
        if "menu_alias" not in ventoy_json:
            ventoy_json["menu_alias"] = []

        for path, new_alias in self.iso_aliases:
            image_path = "/" + path.replace("\\", "/")
            alias_exists = False

            # Determine if the path is a file or directory
            full_path = os.path.join(drive_letter + "\\", path.replace("/", "\\"))
            is_directory = os.path.isdir(full_path)

            for entry in ventoy_json["menu_alias"]:
                key = "dir" if is_directory else "image"
                if entry.get(key) == image_path:
                    entry["alias"] = new_alias
                    alias_exists = True
                    break

            if not alias_exists:
                if is_directory:
                    ventoy_json["menu_alias"].append({"dir": image_path, "alias": new_alias})
                else:
                    ventoy_json["menu_alias"].append({"image": image_path, "alias": new_alias})

        # Save the updated ventoy.json
        self.save_ventoy_json(ventoy_dir, ventoy_json)

        # Clear the list of paths and aliases after applying the rename
        self.iso_aliases.clear()
        self.update_rename_table()


# Run the application
def main():
    app = QtWidgets.QApplication(sys.argv)
    font = QFont("Segoe UI", 9)
    app.setFont(font)

    ventoy_app = VentoyApp()
    ventoy_app.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
