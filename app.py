"""LCD Catalogue Desktop — offline Windows/macOS search interface."""
import json
import sys
from decimal import Decimal
from pathlib import Path
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QSpinBox, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView,
    QPlainTextEdit, QSplitter, QScrollArea,
)
from lookup_engine import lookup
from excel_sync import build_catalogue, save_catalogue_atomic

# In a packaged Windows build, the catalogue stays next to the executable,
# rather than being frozen inside the binary. Updates require no rebuild.
BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
DEFAULT = BASE / 'data' / 'lcd_catalogue.json'
EXCEL_DIR = BASE / 'data' / 'excel'
EXCEL_NAMES = ('COMPATIBLITY LIST', 'PRICE OF LCDS', 'AVAILABLE MODELS HERE WITH ME')


class LCDApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.catalogue = None
        self.quote_items = []
        self.catalogue_path = None
        self.excel_dir = EXCEL_DIR
        self.excel_signature = None
        self.pending_signature = None
        self.setWindowTitle('LCD Catalogue Desktop')
        self.resize(1300, 900)
        self.page_scroll = QScrollArea()
        self.page_scroll.setWidgetResizable(True)
        self.page_scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        self.setCentralWidget(self.page_scroll)

        root = QWidget()
        root.setMinimumWidth(900)
        self.page_scroll.setWidget(root)
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
        refresh = QPushButton('Refresh Catalogue')
        refresh.clicked.connect(self.refresh_and_reset)
        search_bar.addWidget(refresh)
        excel_button = QPushButton('Excel Folder')
        excel_button.clicked.connect(self.choose_excel_folder)
        search_bar.addWidget(excel_button)
        layout.addLayout(search_bar)

        self.catalogue_label = QLabel('No catalogue loaded')
        layout.addWidget(self.catalogue_label)
        self.status_label = QLabel('Enter a phone name or LCD product code to search.')
        self.status_label.setObjectName('status')
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            'LCD Product', 'Sheet', 'Selling Price', 'Store Stock',
            'Availability', 'Review / Action', 'Match Reason'
        ])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        for col in (1, 2, 3, 4):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        # Search results are placed into a resizable splitter below.

        quote_panel = QWidget()
        quote_layout = QVBoxLayout(quote_panel)
        quote_layout.setContentsMargins(0, 4, 0, 0)
        quote_layout.setSpacing(8)

        add_button = QPushButton('Add to Quotation')
        add_button.clicked.connect(self.add_to_quotation)
        add_button.setMaximumWidth(190)
        quote_layout.addWidget(
            add_button, 0, Qt.AlignmentFlag.AlignRight)

        quote_layout.addWidget(QLabel('Current Quotation'))

        self.quote_table = QTableWidget(0, 5)
        self.quote_table.setHorizontalHeaderLabels([
            'LCD Product', 'Quantity', 'Unit Price',
            'Subtotal', 'Availability / Note'
        ])
        self.quote_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers)
        self.quote_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows)
        self.quote_table.setAlternatingRowColors(True)
        self.quote_table.verticalHeader().setVisible(False)
        qh = self.quote_table.horizontalHeader()
        qh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        qh.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        for c in (1, 2, 3):
            qh.setSectionResizeMode(
                c, QHeaderView.ResizeMode.ResizeToContents)
        quote_layout.addWidget(self.quote_table, 1)

        quote_actions = QHBoxLayout()
        self.total_label = QLabel('Total: 0.00')
        quote_layout.addWidget(self.total_label)
        self.total_label.setAlignment(
            Qt.AlignmentFlag.AlignRight)
        self.total_label.setStyleSheet(
            'font-size: 18px; font-weight: bold; padding: 8px;'
        )


        remove_button = QPushButton('Remove Selected')
        remove_button.clicked.connect(self.remove_quote_item)
        quote_actions.addWidget(remove_button)

        copy_button = QPushButton('Copy Quotation')
        copy_button.clicked.connect(self.copy_quotation)
        quote_actions.addWidget(copy_button)

        clear_button = QPushButton('Clear Quotation')
        clear_button.clicked.connect(self.clear_quotation)
        quote_actions.addWidget(clear_button)

        quote_actions.insertStretch(0)
        quote_layout.addLayout(quote_actions)

        self.results_splitter = QSplitter(Qt.Orientation.Vertical)
        self.results_splitter.addWidget(self.table)
        self.results_splitter.addWidget(quote_panel)
        self.results_splitter.setChildrenCollapsible(False)
        self.table.setMinimumHeight(270)
        self.quote_table.setMinimumHeight(290)
        self.results_splitter.setMinimumHeight(650)
        self.results_splitter.setSizes([300, 350])
        layout.addWidget(self.results_splitter, 1)

        self.review_heading = QLabel('Order Review')
        layout.addWidget(self.review_heading)
        self.review = QPlainTextEdit()
        self.review.setReadOnly(True)
        self.review.setFixedHeight(75)
        layout.addWidget(self.review)
        # Position Order Review between search and quotation.
        layout.removeWidget(self.review_heading)
        layout.removeWidget(self.review)
        quote_layout.insertWidget(0, self.review_heading)
        quote_layout.insertWidget(1, self.review)
        quote_layout.removeWidget(add_button)
        quote_layout.insertWidget(
            0, add_button, 0, Qt.AlignmentFlag.AlignRight)


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
        self.excel_timer = QTimer(self)
        self.excel_timer.setInterval(3000)
        self.excel_timer.timeout.connect(self.check_excel_updates)
        self.excel_timer.start()
        self.check_excel_updates()

    def refresh_and_reset(self):
        self.check_excel_updates()
        self.reload_catalogue()
        self.search_input.clear()
        self.qty_input.setValue(1)
        self.table.setRowCount(0)
        self.review.clear()
        self.search_input.setFocus()

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

    def choose_excel_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, 'Choose folder containing the three LCD Excel files', str(self.excel_dir.parent))
        if folder:
            self.excel_dir = Path(folder)
            self.excel_signature = None
            self.pending_signature = None
            self.status_label.setText(f'Watching Excel folder: {self.excel_dir}')
            self.check_excel_updates()

    def excel_files(self):
        if not self.excel_dir.is_dir():
            return None
        files = []
        for prefix in EXCEL_NAMES:
            matches = [p for p in self.excel_dir.iterdir()
                       if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm')
                       and p.name.upper().startswith(prefix)
                       and not p.name.startswith('~$')]
            if len(matches) != 1:
                if len(matches) > 1:
                    raise ValueError(f'Multiple Excel files for {prefix}; keep only one per type')
                return None
            files.append(matches[0])
        return files

    def check_excel_updates(self):
        try:
            files = self.excel_files()
            if not files:
                return
            signature = tuple((str(p), p.stat().st_size, p.stat().st_mtime_ns)
                              for p in files)
            if signature == self.excel_signature:
                return
            if signature != self.pending_signature:
                self.pending_signature = signature
                return  # Wait for another poll, avoiding partially written files.
            catalogue = build_catalogue(*files)
            # Leave previous JSON and in-memory catalogue intact if conversion fails.
            save_catalogue_atomic(catalogue, DEFAULT)
            self.load_catalogue(DEFAULT)
            self.excel_signature = signature
            self.pending_signature = None
            self.status_label.setText('Excel changes synchronized. Search uses updated data.')
        except (OSError, ValueError, PermissionError) as exc:
            self.pending_signature = None
            self.status_label.setText(f'Excel refresh failed; previous catalogue retained: {exc}')
        except Exception as exc:
            self.pending_signature = None
            self.status_label.setText(f'Excel refresh failed; previous catalogue retained: {exc}')

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
                      stock_text, availability, review,
                      item.get('match_reason', 'Not specified')]
            for c, value in enumerate(values):
                cell = QTableWidgetItem(str(value))
                cell.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
                self.table.setItem(r, c, cell)

            self.table.item(r, 0).setData(
                Qt.ItemDataRole.UserRole,
                {
                    'product': item,
                    'requested_model': result['requested_model'],
                    'quantity': self.qty_input.value(),
                    'ambiguous': len(products) > 1,
                    'partial': result.get('partial_search', False),
                }
            )

        label = result['status'].replace('_', ' ')
        self.status_label.setText(
            f'{label}  |  {len(products)} product(s)  |  '
            f'Automatic handling: {"YES" if result["automatic_quote_allowed"] else "NO"}')
        notes = [f'{result["requested_model"]} — {self.qty_input.value()} pcs']
        if result.get('partial_search'):
            notes.append('Partial search: these are candidate LCDs, not confirmed compatible products. Verify the exact phone model.')
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


    def add_to_quotation(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(
                self, 'Select a product',
                'Click the correct LCD in the search results first.')
            return

        cell = self.table.item(row, 0)
        data = cell.data(Qt.ItemDataRole.UserRole) if cell else None

        if not data:
            QMessageBox.information(
                self, 'Select a product',
                'Please select a valid LCD product row.')
            return

        item = data['product']
        price = item.get('main_selling_price')

        if price is None or float(price) <= 0:
            QMessageBox.warning(
                self, 'Missing selling price',
                'This LCD has no confirmed selling price. '
                'It cannot be added to the quotation.')
            return

        if data['ambiguous'] or data['partial']:
            answer = QMessageBox.question(
                self, 'Verify LCD variant',
                'This search has multiple or partial matches.\n\n'
                f'Selected LCD: {item["product"]}\n\n'
                'Have you verified that this is the correct '
                'LCD for the customer?',
                QMessageBox.StandardButton.Yes |
                QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        quantity = data['quantity']
        stock = item['store_quantity']

        if stock <= 0:
            note = 'Big warehouse - order required'
        elif stock < quantity:
            note = 'Local stock insufficient - review'
        else:
            note = 'In store'

        if data['partial']:
            note += '; partial search - verify model'

        self.quote_items.append({
            'product': item['product'],
            'requested': data['requested_model'],
            'quantity': quantity,
            'price': Decimal(str(price)),
            'stock': stock,
            'partial': data['partial'],
            'note': note,
        })

        self.update_quotation()
        self.search_input.clear()
        self.search_input.setFocus()

    def update_quote_quantity(self, row, quantity):
        if not (0 <= row < len(self.quote_items)):
            return

        item = self.quote_items[row]
        item['quantity'] = quantity
        stock = item.get('stock', 0)

        if stock <= 0:
            note = 'Big warehouse - order required'
        elif stock < quantity:
            note = 'Local stock insufficient - review'
        else:
            note = 'In store'

        if item.get('partial'):
            note += '; partial search - verify model'

        item['note'] = note
        self.update_quotation()

    def update_quotation(self):
        self.quote_table.setRowCount(len(self.quote_items))
        total = Decimal('0')

        for r, item in enumerate(self.quote_items):
            subtotal = item['price'] * item['quantity']
            total += subtotal

            self.quote_table.setItem(
                r, 0, QTableWidgetItem(item['product']))

            qty_box = QSpinBox()
            qty_box.setRange(1, 100000)
            qty_box.setValue(item['quantity'])
            qty_box.valueChanged.connect(
                lambda value, row=r:
                    self.update_quote_quantity(row, value))
            self.quote_table.setCellWidget(r, 1, qty_box)

            values = [
                f"{item['price']:.2f}",
                f"{subtotal:.2f}",
                item['note']
            ]
            for c, value in enumerate(values, start=2):
                self.quote_table.setItem(
                    r, c, QTableWidgetItem(value))

        self.total_label.setText(
            f'Total: {total:.2f}  |  '
            f'{len(self.quote_items)} line item(s)')

    def remove_quote_item(self):
        row = self.quote_table.currentRow()
        if 0 <= row < len(self.quote_items):
            self.quote_items.pop(row)
            self.update_quotation()

    def clear_quotation(self):
        if self.quote_items:
            answer = QMessageBox.question(
                self, 'Clear quotation',
                'Remove all items from this quotation?',
                QMessageBox.StandardButton.Yes |
                QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self.quote_items.clear()
        self.update_quotation()
        self.search_input.setFocus()

    def copy_quotation(self):
        if not self.quote_items:
            QMessageBox.information(
                self, 'Empty quotation',
                'Add at least one LCD first.')
            return

        lines = ['LCD QUOTATION', '']
        total = Decimal('0')

        for n, item in enumerate(self.quote_items, 1):
            subtotal = item['price'] * item['quantity']
            total += subtotal

            lines.append(f"{n}. {item['product']}")
            lines.append(
                f"   {item['quantity']} units x "
                f"{item['price']:.2f} = {subtotal:.2f}"
            )

        lines.append(f"TOTAL: {total:.2f}")

        QApplication.clipboard().setText(chr(10).join(lines))

        QMessageBox.information(
            self, 'Quotation copied',
            'Quotation copied to clipboard.')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LCDApp()
    window.show()
    sys.exit(app.exec())
