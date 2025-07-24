# Configure l'environnement pour forcer Matplotlib à utiliser PyQt6
import os

os.environ["QT_API"] = "pyqt6"
import matplotlib

matplotlib.use("QtAgg")

import sys
import re
import json
import sqlite3
import csv
import numpy as np
import pandas as pd
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from scipy import stats
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QComboBox, QCheckBox, QScrollArea, QMessageBox,
    QFormLayout, QDateEdit, QGroupBox, QFrame, QHBoxLayout, QSpacerItem, QSizePolicy,
    QStyle, QDialog, QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QVBoxLayout, QSplitter, QGridLayout
)
from PyQt6.QtCore import QDate, Qt, QRegularExpression
from PyQt6.QtGui import QDoubleValidator, QIntValidator, QFont, QPalette, QColor, QTextCursor, \
    QRegularExpressionValidator
from docx import Document
import sys
import os

# Solution pour PyInstaller
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Désactive Kivy si présent
os.environ['KIVY_NO_CONSOLELOG'] = '1'
class MplCanvas(FigureCanvas):
    """Classe pour intégrer des graphiques matplotlib dans PyQt"""

    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super().__init__(fig)
        self.setParent(parent)


class Variable:
    def __init__(self, nom, description, modalites, type_variable, taille=None):
        self.nom = nom.strip()
        self.description = description.strip()
        self.modalites = self.parse_modalites(modalites)
        self.type_variable = type_variable.strip().upper()
        self.taille = int(taille) if taille and taille.isdigit() else None

    def parse_modalites(self, modalites):
        items = re.findall(r'(\d+)\s*[-:]?\s*([^,\n]+)', modalites)
        return [(int(num.strip()), label.strip()) for num, label in items] if items else []


class AnalysisDialog(QDialog):
    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rapport d'Analyse Exploratoire")
        self.setMinimumSize(1000, 700)

        layout = QVBoxLayout()

        # Onglets pour différents types d'analyse
        self.tabs = QTabWidget()

        # Onglet Résumé
        self.summary_tab = QWidget()
        summary_layout = QVBoxLayout()
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setFont(QFont("Courier New", 10))
        summary_layout.addWidget(self.summary_text)
        self.summary_tab.setLayout(summary_layout)
        self.tabs.addTab(self.summary_tab, "Résumé")

        # Onglet Données Manquantes (seulement le tableau)
        self.missing_tab = QWidget()
        missing_layout = QVBoxLayout()

        # Tableau pour les données manquantes
        self.missing_table = QTableWidget()
        self.missing_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        missing_layout.addWidget(self.missing_table)

        self.missing_tab.setLayout(missing_layout)
        self.tabs.addTab(self.missing_tab, "Données Manquantes")

        # Onglet Distributions Numériques
        self.dist_tab = QWidget()
        dist_layout = QVBoxLayout()

        # Sélecteur de variable numérique
        self.numeric_var_selector = QComboBox()
        dist_layout.addWidget(QLabel("Sélectionnez une variable numérique:"))
        dist_layout.addWidget(self.numeric_var_selector)

        # Conteneur pour les graphiques
        self.dist_graph_container = QWidget()
        self.dist_graph_layout = QGridLayout()
        self.dist_graph_container.setLayout(self.dist_graph_layout)

        dist_layout.addWidget(self.dist_graph_container)
        self.dist_tab.setLayout(dist_layout)
        self.tabs.addTab(self.dist_tab, "Distributions Numériques")

        # Onglet Tests de Normalité
        self.normality_tab = QWidget()
        normality_layout = QVBoxLayout()
        self.normality_table = QTableWidget()
        self.normality_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        normality_layout.addWidget(self.normality_table)
        self.normality_tab.setLayout(normality_layout)
        self.tabs.addTab(self.normality_tab, "Tests de Normalité")

        layout.addWidget(self.tabs)

        # Bouton Fermer
        btn_close = QPushButton("Fermer")
        btn_close.clicked.connect(self.accept)
        btn_close.setStyleSheet("background-color: #4267B2; color: white;")
        layout.addWidget(btn_close)

        self.setLayout(layout)

        # Afficher le rapport
        self.display_report(report)

    def display_report(self, report):
        # Afficher le résumé
        self.summary_text.setPlainText(report['summary'])

        # Afficher les données manquantes dans un tableau
        self.display_missing_table(report['missing_data'])

        # Afficher les distributions numériques
        self.setup_numeric_vars(report['numeric_vars'])

        # Afficher les tests de normalité
        self.display_normality_tests(report['normality_tests'])

    def display_missing_table(self, missing_data):
        """Affiche les données manquantes dans un tableau"""
        if not missing_data:
            return

        # Afficher les données dans un tableau
        self.missing_table.setRowCount(len(missing_data))
        self.missing_table.setColumnCount(4)
        self.missing_table.setHorizontalHeaderLabels(["Variable", "% Manquant", "Type", "Total Manquant"])

        for i, (var, data) in enumerate(missing_data.items()):
            self.missing_table.setItem(i, 0, QTableWidgetItem(var))
            self.missing_table.setItem(i, 1, QTableWidgetItem(f"{data['percentage']:.2f}%"))
            self.missing_table.setItem(i, 2, QTableWidgetItem(data['type']))
            self.missing_table.setItem(i, 3, QTableWidgetItem(str(data['count'])))

        self.missing_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def setup_numeric_vars(self, numeric_vars):
        """Configure le sélecteur de variables numériques"""
        self.numeric_var_selector.clear()
        self.numeric_vars = numeric_vars

        for var in numeric_vars.keys():
            self.numeric_var_selector.addItem(var)

        if numeric_vars:
            self.numeric_var_selector.currentIndexChanged.connect(self.plot_numeric_distribution)
            self.plot_numeric_distribution(0)

    def plot_numeric_distribution(self, index):
        """Affiche les graphiques pour la variable numérique sélectionnée"""
        var_name = self.numeric_var_selector.currentText()
        if not var_name or not self.numeric_vars:
            return

        # Récupérer les données de la variable
        data = self.numeric_vars[var_name]
        values = data['values']

        # Effacer les graphiques précédents
        for i in reversed(range(self.dist_graph_layout.count())):
            widget = self.dist_graph_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # Créer les graphiques
        # Boxplot
        boxplot_canvas = MplCanvas(self, width=5, height=4, dpi=100)
        ax_box = boxplot_canvas.axes
        ax_box.boxplot(values.dropna(), vert=False)
        ax_box.set_title(f'Boxplot de {var_name}')
        ax_box.set_xlabel('Valeurs')
        ax_box.grid(axis='x', linestyle='--', alpha=0.7)

        # Histogramme
        hist_canvas = MplCanvas(self, width=5, height=4, dpi=100)
        ax_hist = hist_canvas.axes
        ax_hist.hist(values.dropna(), bins=20, color='#3498db', edgecolor='black')
        ax_hist.set_title(f'Distribution de {var_name}')
        ax_hist.set_xlabel('Valeurs')
        ax_hist.set_ylabel('Fréquence')
        ax_hist.grid(axis='y', linestyle='--', alpha=0.7)

        # QQ Plot
        qq_canvas = MplCanvas(self, width=5, height=4, dpi=100)
        ax_qq = qq_canvas.axes
        stats.probplot(values.dropna(), dist="norm", plot=ax_qq)
        ax_qq.set_title(f'QQ Plot de {var_name}')
        ax_qq.grid(True, linestyle='--', alpha=0.7)

        # Ajouter les graphiques au layout
        self.dist_graph_layout.addWidget(boxplot_canvas, 0, 0)
        self.dist_graph_layout.addWidget(hist_canvas, 0, 1)
        self.dist_graph_layout.addWidget(qq_canvas, 1, 0, 1, 2)

    def display_normality_tests(self, normality_tests):
        """Affiche les résultats des tests de normalité dans un tableau"""
        if not normality_tests:
            return

        self.normality_table.setRowCount(len(normality_tests))
        self.normality_table.setColumnCount(5)
        self.normality_table.setHorizontalHeaderLabels([
            "Variable", "Shapiro-Wilk (p-value)", "Normalité (Shapiro)",
            "Kolmogorov-Smirnov (p-value)", "Normalité (KS)"
        ])

        for i, (var, tests) in enumerate(normality_tests.items()):
            self.normality_table.setItem(i, 0, QTableWidgetItem(var))

            # Shapiro-Wilk
            shapiro_p = tests.get('shapiro_p', np.nan)
            shapiro_norm = "Oui" if shapiro_p > 0.05 else "Non"
            self.normality_table.setItem(i, 1, QTableWidgetItem(f"{shapiro_p:.4f}"))
            self.normality_table.setItem(i, 2, QTableWidgetItem(shapiro_norm))

            # Kolmogorov-Smirnov
            ks_p = tests.get('ks_p', np.nan)
            ks_norm = "Oui" if ks_p > 0.05 else "Non"
            self.normality_table.setItem(i, 3, QTableWidgetItem(f"{ks_p:.4f}"))
            self.normality_table.setItem(i, 4, QTableWidgetItem(ks_norm))

        self.normality_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


class ExDForm(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ExDForm")
        self.variables = []
        self.inputs = {}
        self.modality_names = {}
        self.db = None
        self.current_db_path = None

        # Appliquer un style global
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 10pt;
            }
            QGroupBox {
                background-color: #ffffff;
                border: 1px solid #dddfe2;
                border-radius: 8px;
                margin-top: 16px;
                padding-top: 10px;
                font-weight: bold;
                color: #1c1e21;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #4267B2;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: 500;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #365899;
            }
            QPushButton:pressed {
                background-color: #29487d;
            }
            QLineEdit, QComboBox, QDateEdit, QCheckBox {
                background-color: #ffffff;
                border: 1px solid #dddfe2;
                border-radius: 4px;
                padding: 6px;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
                border: 1px solid #1877f2;
            }
            QLabel {
                color: #444950;
            }
            QScrollArea {
                border: none;
                background-color: #f0f2f5;
            }
            QFormLayout {
                margin: 10px;
            }
            QMessageBox {
                background-color: #ffffff;
            }
            QDialog {
                background-color: #f0f2f5;
            }
            QTabWidget::pane {
                border: 1px solid #dddfe2;
                background: white;
            }
            QTabBar::tab {
                background: #e4e6eb;
                padding: 8px;
                border: 1px solid #dddfe2;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: white;
                margin-bottom: -1px;
            }
            QTableWidget {
                gridline-color: #dddfe2;
                selection-background-color: #e7f3ff;
            }
            QHeaderView::section {
                background-color: #4267B2;
                color: white;
                padding: 4px;
                border: 1px solid #3a5a99;
            }
            QComboBox {
                background-color: white;
            }
        """)

        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Titre de l'application
        title_label = QLabel("Formulaire de saisie de Données")
        title_label.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #1c1e21; margin-bottom: 20px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        # Section Gestion des bases de données
        db_group = QGroupBox("Gestion des Bases de Données")
        db_layout = QVBoxLayout()
        db_layout.setSpacing(10)

        self.current_db_label = QLabel("Base de données actuelle : Aucune")
        self.current_db_label.setStyleSheet("color: #606770; font-style: italic;")

        db_buttons_layout = QHBoxLayout()
        db_buttons_layout.setSpacing(10)

        btn_new_db = QPushButton("Nouvelle base")
        btn_new_db.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        btn_new_db.clicked.connect(self.new_database)

        btn_open_db = QPushButton("Ouvrir une base")
        btn_open_db.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_open_db.clicked.connect(self.open_database)

        btn_reset_db = QPushButton("Réinitialiser la base")
        btn_reset_db.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        btn_reset_db.clicked.connect(self.reset_database)
        btn_reset_db.setStyleSheet("background-color: #e74c3c;")

        db_buttons_layout.addWidget(btn_new_db)
        db_buttons_layout.addWidget(btn_open_db)
        db_buttons_layout.addWidget(btn_reset_db)

        db_layout.addWidget(self.current_db_label)
        db_layout.addLayout(db_buttons_layout)
        db_group.setLayout(db_layout)
        main_layout.addWidget(db_group)

        # Section Fonctionnalités principales
        func_group = QGroupBox("Fonctionnalités")
        func_layout = QVBoxLayout()
        func_layout.setSpacing(10)

        btn_import = QPushButton("Importer tableau de variables")
        btn_import.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView))
        btn_import.clicked.connect(self.import_docx)

        btn_save = QPushButton("Enregistrer dans la base")
        btn_save.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        btn_save.clicked.connect(self.save_data)
        btn_save.setStyleSheet("background-color: #27ae60;")

        btn_export = QPushButton("Exporter en CSV")
        btn_export.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowDown))
        btn_export.clicked.connect(self.export_csv)
        btn_export.setStyleSheet("background-color: #2c3e50;")

        # Nouveau bouton pour l'analyse exploratoire
        btn_analysis = QPushButton("Analyse Exploratoire")
        btn_analysis.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        btn_analysis.clicked.connect(self.show_exploratory_analysis)
        btn_analysis.setStyleSheet("background-color: #9b59b6;")

        func_buttons_layout = QHBoxLayout()
        func_buttons_layout.addWidget(btn_import)
        func_buttons_layout.addWidget(btn_save)
        func_buttons_layout.addWidget(btn_export)
        func_buttons_layout.addWidget(btn_analysis)

        func_layout.addLayout(func_buttons_layout)
        func_group.setLayout(func_layout)
        main_layout.addWidget(func_group)

        # Section Formulaire
        form_group = QGroupBox("Formulaire de Saisie")
        form_layout = QVBoxLayout()
        form_layout.setSpacing(15)

        self.form_area = QScrollArea()
        self.form_area.setWidgetResizable(True)
        self.form_area.setFrameShape(QFrame.Shape.NoFrame)

        self.form_container = QWidget()
        self.form_layout = QFormLayout()
        self.form_layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        self.form_layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        self.form_layout.setVerticalSpacing(15)
        self.form_layout.setHorizontalSpacing(20)

        self.form_container.setLayout(self.form_layout)
        self.form_area.setWidget(self.form_container)

        form_layout.addWidget(self.form_area)
        form_group.setLayout(form_layout)
        main_layout.addWidget(form_group, 1)

        # Statut en bas de fenêtre
        status_bar = QFrame()
        status_bar.setFrameShape(QFrame.Shape.StyledPanel)
        status_bar.setStyleSheet("background-color: #4267B2; color: white; padding: 5px;")
        status_layout = QHBoxLayout()

        self.status_label = QLabel("Prêt")
        self.status_label.setStyleSheet("color: white;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignLeft)

        status_layout.addWidget(self.status_label)
        status_bar.setLayout(status_layout)
        main_layout.addWidget(status_bar)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def update_status(self, message):
        self.status_label.setText(message)
        QApplication.processEvents()

    def connect_to_database(self, path):
        if self.db:
            self.db.close()
        self.db = sqlite3.connect(path)
        self.current_db_path = path
        self.create_table()
        self.current_db_label.setText(f"Base de données actuelle : {self.current_db_path}")
        self.update_status(f"Base de données connectée: {path}")

    def new_database(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Nouvelle base de données",
            "",
            "Fichiers SQLite (*.db *.sqlite)"
        )
        if file_path:
            if not file_path.lower().endswith(('.db', '.sqlite')):
                file_path += '.db'
            self.connect_to_database(file_path)
            QMessageBox.information(self, "Succès", f"Nouvelle base de données créée : {file_path}")
            self.update_status(f"Base créée: {file_path}")

    def open_database(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Ouvrir une base de données",
            "",
            "Fichiers SQLite (*.db *.sqlite)"
        )
        if file_path:
            self.connect_to_database(file_path)
            QMessageBox.information(self, "Succès", f"Base de données ouverte : {file_path}")
            self.update_status(f"Base ouverte: {file_path}")

    def reset_database(self):
        if self.db:
            reply = QMessageBox.question(
                self,
                'Confirmer',
                "Voulez-vous vraiment réinitialiser la base de données? Toutes les données seront perdues.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                cursor = self.db.cursor()
                cursor.execute("DELETE FROM data")
                self.db.commit()
                QMessageBox.information(self, "Succès", "Base de données réinitialisée.")
                self.update_status("Base réinitialisée")
        else:
            QMessageBox.warning(self, "Avertissement", "Aucune base de données n'est ouverte.")
            self.update_status("Erreur: aucune base ouverte")

    def create_table(self):
        if self.db:
            cursor = self.db.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    form_data TEXT
                )
            """)
            self.db.commit()

    def import_docx(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Choisir un fichier Word", "", "Documents Word (*.docx)")
        if file_path:
            self.variables = self.read_variables_from_docx(file_path)
            self.generate_form()
            self.update_status(f"Fichier importé: {file_path.split('/')[-1]}")

    def read_variables_from_docx(self, file_path):
        doc = Document(file_path)
        vars = []
        for table in doc.tables:
            # Vérifier si la table a au moins 5 colonnes (nouveau format)
            if len(table.rows) > 0 and len(table.rows[0].cells) >= 5:
                for row in table.rows[1:]:
                    cells = row.cells
                    if len(cells) >= 5:
                        nom = cells[0].text
                        description = cells[1].text
                        modalites = cells[2].text
                        type_variable = cells[3].text
                        taille = cells[4].text  # Nouvelle colonne pour la taille

                        # Handle multi-line modalities by replacing newlines with commas
                        modalites = modalites.replace('\n', ', ')

                        if type_variable.strip().upper() != "ID":
                            vars.append(Variable(nom, description, modalites, type_variable, taille))
            else:
                # Ancien format sans colonne de taille
                for row in table.rows[1:]:
                    cells = row.cells
                    if len(cells) >= 4:
                        nom = cells[0].text
                        description = cells[1].text
                        modalites = cells[2].text
                        type_variable = cells[3].text

                        # Handle multi-line modalities by replacing newlines with commas
                        modalites = modalites.replace('\n', ', ')

                        if type_variable.strip().upper() != "ID":
                            vars.append(Variable(nom, description, modalites, type_variable))
        return vars

    def generate_form(self):
        # Supprimer les anciens widgets
        while self.form_layout.count():
            child = self.form_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.inputs.clear()
        self.modality_names = {}

        if not self.variables:
            no_vars_label = QLabel("Aucune variable importée. Veuillez importer un fichier Word.")
            no_vars_label.setStyleSheet("color: #7f8c8d; font-style: italic;")
            no_vars_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.form_layout.addRow(no_vars_label)
            return

        for var in self.variables:
            group = QGroupBox(f"{var.nom} - {var.description}")
            group_layout = QFormLayout()
            group_layout.setVerticalSpacing(10)
            group_layout.setHorizontalSpacing(15)

            if var.type_variable == "NUM_CONTINUE":
                input_field = QLineEdit()
                # Validateur pour nombres à virgule flottante (point décimal)
                validator = QDoubleValidator()
                validator.setNotation(QDoubleValidator.Notation.StandardNotation)
                input_field.setValidator(validator)
                input_field.setPlaceholderText("Entrez un nombre décimal (ex: 12.34)")

                # Ajouter le contrôle de taille si disponible
                if var.taille:
                    input_field.setMaxLength(var.taille)

                self.inputs[var.nom] = input_field
                group_layout.addRow(QLabel("Valeur:"), input_field)

            elif var.type_variable == "NUM_DISCRETE":
                input_field = QLineEdit()
                # Validateur pour nombres entiers
                validator = QIntValidator()
                input_field.setValidator(validator)
                input_field.setPlaceholderText("Entrez un nombre entier")

                # Ajouter le contrôle de taille si disponible
                if var.taille:
                    input_field.setMaxLength(var.taille)

                self.inputs[var.nom] = input_field
                group_layout.addRow(QLabel("Valeur:"), input_field)

            elif var.type_variable == "TEXTE":
                input_field = QLineEdit()
                input_field.setPlaceholderText("Entrez du texte")

                # Ajouter le contrôle de taille si disponible
                if var.taille:
                    input_field.setMaxLength(var.taille)

                self.inputs[var.nom] = input_field
                group_layout.addRow(QLabel("Valeur:"), input_field)

            elif var.type_variable == "BINAIRE":
                combo = QComboBox()
                combo.addItems(["1 (Oui)", "0 (Non)"])
                self.inputs[var.nom] = combo
                group_layout.addRow(QLabel("Sélection:"), combo)

            elif var.type_variable == "CATEGORIELLE":
                combo = QComboBox()
                combo.addItem("-- Sélectionnez une option --", "")
                for num, label_mod in var.modalites:
                    combo.addItem(f"{num} - {label_mod}", num)
                self.inputs[var.nom] = combo
                group_layout.addRow(QLabel("Sélection:"), combo)

            elif var.type_variable == "CATEGORIELLE_MULTIPLE":
                # MODIFICATION: Utiliser des combobox au lieu de checkbox
                for num, mod in var.modalites:
                    combo = QComboBox()
                    combo.addItem("-- Sélectionnez --", "")  # Valeur vide par défaut
                    combo.addItem("1 (Oui)", 1)
                    combo.addItem("0 (Non)", 0)

                    unique_id = f"{var.nom}_{num}"
                    self.modality_names[unique_id] = mod
                    if var.nom not in self.inputs:
                        self.inputs[var.nom] = []
                    self.inputs[var.nom].append((unique_id, combo))
                    group_layout.addRow(QLabel(f"{mod}:"), combo)

            elif var.type_variable == "DATE":
                date_edit = QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDate(QDate.currentDate())
                date_edit.setDisplayFormat("dd/MM/yyyy")
                self.inputs[var.nom] = date_edit
                group_layout.addRow(QLabel("Date:"), date_edit)

            elif var.type_variable == "TEMPS":
                time_field = QLineEdit()
                # Validateur pour le format hh:mm:ss
                regex = QRegularExpression("^([0-1][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$")
                validator = QRegularExpressionValidator(regex)
                time_field.setValidator(validator)
                time_field.setPlaceholderText("hh:mm:ss")
                self.inputs[var.nom] = time_field
                group_layout.addRow(QLabel("Heure:"), time_field)

            group.setLayout(group_layout)
            self.form_layout.addRow(group)

        # Installer un event filter pour la navigation avec Entrée
        for widget in self.findChildren(QLineEdit) + self.findChildren(QComboBox) + self.findChildren(QDateEdit):
            widget.installEventFilter(self)

    def eventFilter(self, obj, event):
        # Seulement pour les événements clavier sur les widgets d'entrée
        if (event.type() == event.Type.KeyPress and
                event.key() in [Qt.Key.Key_Return, Qt.Key.Key_Enter]):

            # Pour les combobox, ne pas avancer si le menu déroulant est ouvert
            if isinstance(obj, QComboBox) and obj.view().isVisible():
                return False

            # Empêcher la propagation de l'événement Entrée
            event.accept()

            # Passer au widget suivant
            next_widget = self.focusNextChild()
            if next_widget:
                next_widget.setFocus()
            return True

        return super().eventFilter(obj, event)

    def save_data(self):
        if not self.db:
            QMessageBox.warning(self, "Erreur", "Veuillez d'abord créer ou ouvrir une base de données.")
            self.update_status("Erreur: aucune base de données ouverte")
            return

        data = {}
        error_fields = []
        self.update_status("Validation des données...")

        for var in self.variables:
            if var.type_variable == "NUM_CONTINUE":
                field = self.inputs[var.nom]
                value = field.text().strip()

                if value:
                    # Vérifier le séparateur décimal
                    if ',' in value:
                        error_fields.append(f"{var.nom}: Utilisez le point (.) comme séparateur décimal")
                        field.setStyleSheet("border: 1px solid red;")
                    else:
                        try:
                            # Convertir en float pour vérifier
                            float_value = float(value)
                            data[var.nom] = value
                            field.setStyleSheet("")
                        except ValueError:
                            error_fields.append(f"{var.nom}: Valeur numérique invalide")
                            field.setStyleSheet("border: 1px solid red;")
                else:
                    # Champ vide accepté
                    data[var.nom] = value
                    field.setStyleSheet("")

            elif var.type_variable == "NUM_DISCRETE":
                field = self.inputs[var.nom]
                value = field.text().strip()
                try:
                    if value:  # Allow empty fields
                        int(value)
                    data[var.nom] = value
                    field.setStyleSheet("")
                except ValueError:
                    error_fields.append(f"{var.nom}: Doit être un nombre entier")
                    field.setStyleSheet("border: 1px solid red;")

            elif var.type_variable == "TEXTE":
                data[var.nom] = self.inputs[var.nom].text().strip()

            elif var.type_variable == "BINAIRE":
                data[var.nom] = 1 if "1" in self.inputs[var.nom].currentText() else 0

            elif var.type_variable == "CATEGORIELLE":
                data[var.nom] = self.inputs[var.nom].currentData()

            elif var.type_variable == "DATE":
                data[var.nom] = self.inputs[var.nom].date().toString("yyyy-MM-dd")

            elif var.type_variable == "TEMPS":
                data[var.nom] = self.inputs[var.nom].text().strip()

            elif var.type_variable == "CATEGORIELLE_MULTIPLE":
                for unique_id, combobox in self.inputs[var.nom]:
                    modalite_name = self.modality_names[unique_id]
                    # Récupérer la valeur sélectionnée (peut être vide, 1 ou 0)
                    value = combobox.currentData()
                    if value != "":
                        data[modalite_name] = value
                    # Si rien n'est sélectionné, la clé n'est pas ajoutée (reste vide dans la base)

        # Vérifier s'il y a des erreurs
        if error_fields:
            QMessageBox.warning(
                self,
                "Erreurs de saisie",
                "Veuillez corriger les champs suivants :\n\n" + "\n".join(error_fields))
            self.update_status("Erreurs dans le formulaire")
            return

        try:
            cursor = self.db.cursor()
            cursor.execute("INSERT INTO data (form_data) VALUES (?)", (json.dumps(data),))
            self.db.commit()
            QMessageBox.information(self, "Succès", "Données enregistrées avec succès.")
            self.generate_form()  # Réinitialiser le formulaire
            self.update_status("Données enregistrées avec succès")
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Erreur", f"Erreur de base de données : {str(e)}")
            self.update_status("Erreur lors de l'enregistrement")

    def export_csv(self):
        if not self.db:
            QMessageBox.warning(self, "Erreur", "Aucune base de données ouverte.")
            self.update_status("Erreur: aucune base de données ouverte")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Exporter en CSV", "", "CSV Files (*.csv)")
        if file_path:
            try:
                self.update_status("Exportation en cours...")
                cursor = self.db.cursor()
                cursor.execute("SELECT id, form_data FROM data")
                rows = cursor.fetchall()

                all_data = []
                for row in rows:
                    d = json.loads(row[1])
                    d["id"] = row[0]
                    all_data.append(d)

                if not all_data:
                    QMessageBox.warning(self, "Avertissement", "Aucune donnée à exporter.")
                    self.update_status("Aucune donnée à exporter")
                    return

                # Créer une liste ordonnée des clés basée sur l'ordre des variables
                ordered_keys = ["id"]

                # Ajouter les variables dans l'ordre de la table
                for var in self.variables:
                    if var.type_variable == "CATEGORIELLE_MULTIPLE":
                        for num, mod in var.modalites:
                            if mod in all_data[0]:
                                ordered_keys.append(mod)
                    else:
                        if var.nom in all_data[0]:
                            ordered_keys.append(var.nom)

                # Ajouter les autres clés qui pourraient manquer
                all_keys = set().union(*(d.keys() for d in all_data))
                missing_keys = [k for k in all_keys if k not in ordered_keys and k != "id"]
                ordered_keys.extend(sorted(missing_keys))

                with open(file_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.DictWriter(f, fieldnames=ordered_keys)
                    writer.writeheader()
                    for row in all_data:
                        # S'assurer que toutes les clés sont présentes
                        writer.writerow({k: row.get(k, "") for k in ordered_keys})

                QMessageBox.information(self, "Succès", f"Exporté vers {file_path}")
                self.update_status(f"Export réussi: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Erreur lors de l'export : {str(e)}")
                self.update_status("Erreur lors de l'export")

    def show_exploratory_analysis(self):
        """Affiche le rapport d'analyse exploratoire"""
        if not self.db:
            QMessageBox.warning(self, "Erreur", "Aucune base de données ouverte.")
            self.update_status("Erreur: aucune base de données ouverte")
            return

        self.update_status("Génération du rapport d'analyse...")
        report = self.generate_analysis_report()
        self.update_status("Rapport généré")

        # Afficher le rapport dans une fenêtre modale
        if "error" not in report:
            analysis_dialog = AnalysisDialog(report, self)
            analysis_dialog.exec()
        else:
            QMessageBox.warning(self, "Erreur", report["error"])

    def generate_analysis_report(self):
        """Génère un rapport d'analyse exploratoire"""
        if not self.db:
            return {"error": "Aucune base de données ouverte"}

        cursor = self.db.cursor()
        cursor.execute("SELECT id, form_data FROM data")
        rows = cursor.fetchall()

        if not rows:
            return {"error": "Aucune donnée disponible pour l'analyse"}

        # Convertir les données en DataFrame pandas pour faciliter l'analyse
        data_list = []
        for row in rows:
            record = json.loads(row[1])
            record['id'] = row[0]
            data_list.append(record)

        df = pd.DataFrame(data_list)

        # 1. Rapport sur les données manquantes
        missing_data = {}
        total_records = len(df)

        for col in df.columns:
            if col == 'id':
                continue

            missing_count = df[col].isna().sum() if col in df.columns else total_records
            missing_percentage = (missing_count / total_records) * 100

            # Déterminer le type de variable
            var_type = "Inconnu"
            for var in self.variables:
                if col == var.nom or col in [self.modality_names.get(f"{var.nom}_{num}") for num, mod in var.modalites]:
                    var_type = var.type_variable
                    break

            missing_data[col] = {
                "count": int(missing_count),
                "percentage": missing_percentage,
                "type": var_type
            }

        # 2. Variables numériques pour les distributions
        numeric_vars = {}
        for var in self.variables:
            if var.type_variable in ["NUM_CONTINUE", "NUM_DISCRETE"] and var.nom in df.columns:
                # Convertir en numérique
                df[var.nom] = pd.to_numeric(df[var.nom], errors='coerce')
                numeric_vars[var.nom] = {
                    "values": df[var.nom],
                    "type": var.type_variable
                }

        # 3. Tests de normalité pour les variables numériques
        normality_tests = {}
        for var_name, data in numeric_vars.items():
            values = data['values'].dropna()

            if len(values) > 3:  # Minimum 3 valeurs pour les tests
                # Test Shapiro-Wilk
                shapiro_test = stats.shapiro(values)

                # Test Kolmogorov-Smirnov
                ks_test = stats.kstest(values, 'norm', args=(values.mean(), values.std()))

                normality_tests[var_name] = {
                    "shapiro_p": shapiro_test.pvalue,
                    "ks_p": ks_test.pvalue
                }

        # 4. Générer un résumé textuel
        summary = "Rapport d'Analyse Exploratoire\n"
        summary += "=" * 40 + "\n\n"
        summary += f"Nombre total d'enregistrements: {total_records}\n"
        summary += f"Nombre de variables: {len(df.columns) - 1}\n\n"

        summary += "Variables Numériques:\n"
        summary += "-" * 40 + "\n"
        for var_name in numeric_vars:
            n_missing = df[var_name].isna().sum()
            p_missing = (n_missing / total_records) * 100
            summary += f"{var_name}: {len(df[var_name].dropna())} valeurs, {p_missing:.1f}% manquants\n"
        summary += "\n"

        summary += "Variables Catégorielles:\n"
        summary += "-" * 40 + "\n"
        for var in self.variables:
            if var.type_variable in ["CATEGORIELLE", "BINAIRE", "CATEGORIELLE_MULTIPLE"]:
                n_missing = df[var.nom].isna().sum() if var.nom in df.columns else total_records
                p_missing = (n_missing / total_records) * 100
                summary += f"{var.nom}: {p_missing:.1f}% manquants\n"

        # Préparer le rapport final
        report = {
            "summary": summary,
            "missing_data": missing_data,
            "numeric_vars": numeric_vars,
            "normality_tests": normality_tests,
            "total_records": total_records
        }

        return report


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExDForm()
    window.show()
    sys.exit(app.exec())

    # Définir une palette de couleurs cohérente
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(240, 242, 245))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(28, 30, 33))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 242, 245))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(28, 30, 33))
    palette.setColor(QPalette.ColorRole.Text, QColor(28, 30, 33))
    palette.setColor(QPalette.ColorRole.Button, QColor(66, 103, 178))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(66, 103, 178))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    window = ExDForm()
    window.resize(1000, 700)
    window.show()
    sys.exit(app.exec())