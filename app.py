"""LCD Catalogue Desktop — offline Windows/macOS search interface."""
import json
import sys
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QSpinBox, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView,
    QPlainTextEdit,
)
from lookup_engine import lookup

# In a packaged Windows build, the catalogue stays next to the executable,
# rather than being frozen inside the binary. Updates require no rebuild.
BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
DEFAULT = BASE / 'data' / 'lcd_catalogue.json'


class LCDApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.catalogue = None
        self.catalogue_path = None
        self.setWindowTitle('LCD Catalogue Desktop')
        self.resize(1150, 760)
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        heading = QLabel('LCD Catalogue')
        heading.setObjectName('heading')
        layout.addWidget(heading)
        layout.addWidget(QLabel('Offline compatibility, pricing and local inventory'))

        search_bar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Phone name or LCD code — e.g. Realme 10 5G')
        self.search_input.returnPressed.connect(self.search)
        search_bar.addWidget(self.search_input, 1)
        search_bar.addWidget(QLabel('Quantity:'))
        self.qty_input = QSpinBox()
        self.qty_input.setRange(1, 100000)
        self.qty_input.setValue(1)
        self.qty_input.setFixedWidth(100)
        search_bar.addWidget(self.qty_input)
        go = QPushButton('Search')
        go.clicked.connect(self.search)
        search_bar.addWidget(go)
        load = QPushButton('Load Catalogue')
        load.clicked.connect(self.choose_catalogue)
        search_bar.addWidget(load)
        refresh = QPushButton('Reload Catalogue')
        refresh.clicked.connect(self.reload_catalogue)
        search_bar.addWidget(refresh)
        layout.addLayout(search_bar)

        self.catalogue_label = QLabel('No catalogue loaded')
        layout.addWidget(self.catalogue_label)
        self.status_label = QLabel('Enter a phone name or LCD product code to search.')
        self.status_label.setObjectName('status')
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels([
            'LCD Product', 'Sheet', 'Selling Price', 'Store Stock',
            'Availability', 'Review / Action'
        ])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        for col in (1, 2, 3, 4):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        layout.addWidget(QLabel('Order review'))
        self.review = QPlainTextEdit()
        self.review.setReadOnly(True)
        self.review.setMaximumHeight(125)
        layout.addWidget(self.review)
        self.note = QLabel('Selling prices and stock are from the loaded local catalogue. '
                           'Products listed in the price catalogue can be ordered from the big warehouse.')
        self.note.setWordWrap(True)
        layout.addWidget(self.note)
        self.setStyleSheet('''
          QWidget { background:#f8fafc; color:#172033; font-size:13px; }
          QLabel#heading { font-size:26px; font-weight:bold; }
          QLabel#status { font-weight:bold; padding:8px; background:#e8eef9; border-radius:5px; }
          QLineEdit, QSpinBox, QPlainTextEdit { background:white; padding:9px;
              border:1px solid #cbd5e1; border-radius:5px; }
          QPushButton { background:#174ea6; color:white; padding:10px 13px;
              border:0; border-radius:5px; }
          QPushButton:hover { background:#123c80; }
          QTableWidget { background:white; alternate-background-color:#f1f5f9;
              gridline-color:#e2e8f0; }
          QHeaderView::section { background:#e9eef5; padding:7px; font-weight:bold; }
        ''')
        if DEFAULT.is_file():
            self.load_catalogue(DEFAULT)

    def load_catalogue(self, path):
        try:
            with path.open(encoding='utf-8') as fh:
                catalogue = json.load(fh)
            for key in ('compatibility', 'prices', 'inventory'):
                if not isinstance(catalogue.get(key), list):
                    raise ValueError(f'Invalid catalogue: missing {key} array')
            self.catalogue = catalogue
            self.catalogue_path = path
            self.catalogue_label.setText(
                f'Catalogue: {len(catalogue["compatibility"])} compatibility records  |  '
                f'{len(catalogue["prices"])} products  |  '
                f'{len(catalogue["inventory"])} inventory entries  |  '
                f'Updated: {catalogue.get("generated_at", "unknown")}')
            self.table.setRowCount(0)
            self.review.clear()
            self.status_label.setText('Catalogue loaded. Ready to search.')
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            QMessageBox.critical(self, 'Catalogue error', str(exc))

    def choose_catalogue(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, 'Select LCD catalogue', str(BASE), 'JSON files (*.json)')
        if filename:
            self.load_catalogue(Path(filename))

    def reload_catalogue(self):
        if not self.catalogue_path:
            QMessageBox.information(self, 'No catalogue', 'Load a JSON catalogue first.')
            return
        self.load_catalogue(self.catalogue_path)

    def search(self):
        if self.catalogue is None:
            QMessageBox.warning(self, 'Missing catalogue', 'Load a catalogue first.')
            return
        try:
            result = lookup(self.catalogue, self.search_input.text(), self.qty_input.value())
        except ValueError as exc:
            self.status_label.setText(str(exc))
            return
        products = result['products']
        self.table.setRowCount(len(products))
        for r, item in enumerate(products):
            price = item['main_selling_price']
            qty = item['store_quantity']
            reasons = item['review_reasons']
            price_text = 'Confirm price' if price is None else f'{price:.2f}'
            stock_text = str(int(qty)) if float(qty).is_integer() else str(qty)
            availability = 'In store' if qty > 0 else 'Big warehouse - available to order'
            if len(products) > 1:
                review = 'Confirm correct LCD variant'
            elif reasons:
                review = ', '.join(s.replace('_', ' ').capitalize() for s in reasons)
            else:
                review = 'No review required'
            values = [item['product'], item['source_sheet'], price_text,
                      stock_text, availability, review]
            for c, value in enumerate(values):
                cell = QTableWidgetItem(str(value))
                cell.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
                self.table.setItem(r, c, cell)

        label = result['status'].replace('_', ' ')
        self.status_label.setText(
            f'{label}  |  {len(products)} product(s)  |  '
            f'Automatic handling: {"YES" if result["automatic_quote_allowed"] else "NO"}')
        notes = [f'{result["requested_model"]} — {self.qty_input.value()} pcs']
        if not products:
            notes.append('No matching LCD in the loaded price catalogue. Manual verification required.')
        if len(products) > 1:
            notes.append('Multiple LCD alternatives: confirm the exact phone/display variant before ordering.')
        if any(not p['price_confirmed'] for p in products):
            notes.append('One or more products have no confirmed selling price. Do not use purchase price.')
        if any(p['warehouse_confirmation_required'] for p in products):
            notes.append('Big warehouse order or local stock shortage requires action.')
        if result['automatic_quote_allowed']:
            notes.append('One matched product; catalogue price and local quantity are sufficient.')
        if products and not result['automatic_quote_allowed']:
            notes.append('Manual review required before confirming the order.')
        self.review.setPlainText('\n'.join(notes))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LCDApp()
    window.show()
    sys.exit(app.exec())
