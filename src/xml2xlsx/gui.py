# Released under GPL3 License.
# Copyright (c) 2026 Ladislav Bartos and Robert Vacha Lab

import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Self

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QDragEnterEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ._version import __version__
from .core import Summary, convert


@dataclass(frozen=True)
class Palette:
    """Colours of the window. One place to change if it should look different."""

    background: str = "#f6f5f4"
    surface: str = "#ffffff"
    surface_hot: str = "#eef4fd"
    border: str = "#d8d4d0"
    heading: str = "#171a1f"
    text: str = "#3d4148"
    muted: str = "#8b8e95"
    accent: str = "#3584e4"
    accent_hot: str = "#2b74cc"
    accent_down: str = "#1c64bd"
    accent_off: str = "#c8cbd0"
    on_accent: str = "#ffffff"


PALETTE = Palette()

# how many file names fit in the drop zone before it starts counting
NAMES_SHOWN = 3


class Outcome(Enum):
    """
    What happened to one plate, and the colour it is shown in.

    The report's status column starts with one of these words, so the
    prefix is both the member value and what we match on.
    """

    prefix: str
    colour: str

    WRITTEN = ("written", "#26794a")
    SKIPPED = ("skipped", "#8b8e95")
    WARNING = ("warning", "#a15c00")
    ERROR = ("error", "#c01c28")

    def __new__(cls, prefix: str, colour: str) -> Self:
        outcome = object.__new__(cls)
        outcome._value_ = prefix
        outcome.prefix = prefix
        outcome.colour = colour

        return outcome

    @classmethod
    def of(cls, status: str) -> "Outcome":
        """The outcome a status line describes. Anything unexpected reads as skipped."""
        for outcome in cls:
            if status.startswith(outcome.prefix):
                return outcome

        return cls.SKIPPED


@dataclass(frozen=True)
class Column:
    """
    One column of the report table.

    Attributes:
        label: Heading shown to the user.
        index: Position of this field in a report row.
        stretch: Whether it takes the slack when the window is resized.
    """

    label: str
    index: int
    stretch: bool = False


COLUMNS = (
    Column("Sheet", 0),
    Column("Plate", 1),
    Column("Experiment", 2),
    Column("Read time", 3),
    Column("Source", 4),
    Column("Status", 5, stretch=True),
)


def stylesheet(palette: Palette) -> str:
    """The whole look of the window, as Qt's flavour of CSS."""
    return f"""
    QWidget {{
        background: {palette.background};
        color: {palette.text};
        font-size: 14px;
    }}
    QLabel#heading {{
        font-size: 22px;
        font-weight: 600;
        color: {palette.heading};
    }}
    QLabel#status {{ color: {palette.text}; }}

    QFrame#zone {{
        background: {palette.surface};
        border: 2px solid {palette.border};
        border-radius: 14px;
    }}
    QFrame#zone[hot="true"] {{
        background: {palette.surface_hot};
        border-color: {palette.accent};
    }}
    /* without this the labels paint a square over the rounded corners */
    QFrame#zone QLabel {{ background: transparent; border: none; }}
    QLabel#zoneTitle {{
        font-size: 16px;
        font-weight: 600;
        color: {palette.heading};
    }}
    QLabel#zoneSubtitle {{ color: {palette.muted}; }}

    QPushButton#convert {{
        background: {palette.accent};
        color: {palette.on_accent};
        border: none;
        border-radius: 8px;
        padding: 12px 26px;
        font-weight: 600;
    }}
    QPushButton#convert:hover {{ background: {palette.accent_hot}; }}
    QPushButton#convert:pressed {{ background: {palette.accent_down}; }}
    QPushButton#convert:disabled {{ background: {palette.accent_off}; }}

    QPushButton#clear {{
        background: transparent;
        color: {palette.muted};
        border: none;
        border-radius: 8px;
        padding: 12px 20px;
    }}
    QPushButton#clear:hover {{
        background: {palette.surface_hot};
        color: {palette.accent};
    }}

    QProgressBar {{
        background: {palette.border};
        border: none;
        border-radius: 2px;
        max-height: 4px;
    }}
    QProgressBar::chunk {{
        background: {palette.accent};
        border-radius: 2px;
    }}

    QTableWidget {{
        background: {palette.surface};
        border: 1px solid {palette.border};
        border-radius: 10px;
        gridline-color: transparent;
    }}
    QTableWidget::item {{ padding: 6px 4px; border: none; }}
    QHeaderView::section {{
        background: {palette.background};
        color: {palette.muted};
        border: none;
        border-bottom: 1px solid {palette.border};
        padding: 9px 4px;
        font-weight: 600;
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 10px;
        margin: 4px;
    }}
    QScrollBar::handle:vertical {{
        background: {palette.border};
        border-radius: 3px;
        min-height: 30px;
    }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
    """


def paths_from(event: QDropEvent) -> list[Path]:
    """The local files a drop carried, ignoring anything that is not one."""
    return [
        Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()
    ]


class DropZone(QFrame):
    """
    The large panel files are dropped onto, or clicked to browse.

    Attributes:
        dropped: Emitted with the paths of files dropped on the panel.
        clicked: Emitted when the panel is clicked.
    """

    dropped = pyqtSignal(list)
    clicked = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("zone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setProperty("hot", False)

        self.title = QLabel("Drop XML files here")
        self.title.setObjectName("zoneTitle")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.subtitle = QLabel("or click to choose them")
        self.subtitle.setObjectName("zoneSubtitle")
        self.subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.subtitle.setWordWrap(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.addStretch()
        layout.addWidget(self.title)
        layout.addWidget(self.subtitle)
        layout.addStretch()

    def show_text(self, title: str, subtitle: str) -> None:
        """Replace the two lines of text in the panel."""
        self.title.setText(title)
        self.subtitle.setText(subtitle)

    def _set_hot(self, hot: bool) -> None:
        """Light the panel up, or stop. Qt only restyles when asked to."""
        self.setProperty("hot", hot)
        style = self.style()
        style.unpolish(self)
        style.polish(self)

    def enterEvent(self, event) -> None:
        self._set_hot(True)

    def leaveEvent(self, a0) -> None:
        self._set_hot(False)

    def mousePressEvent(self, a0: QMouseEvent) -> None:
        if a0.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

    def dragEnterEvent(self, a0: QDragEnterEvent) -> None:
        if a0.mimeData().hasUrls():
            a0.acceptProposedAction()
            self._set_hot(True)

    def dragLeaveEvent(self, a0) -> None:
        self._set_hot(False)

    def dropEvent(self, a0: QDropEvent) -> None:
        self._set_hot(False)
        self.dropped.emit(paths_from(a0))
        a0.acceptProposedAction()


class Conversion(QObject):
    """
    The conversion, run on a worker thread.

    Attributes:
        done: Emitted with the report, or with the exception that stopped it.
    """

    done = pyqtSignal(object)

    def __init__(self, sources: list[Path], destination: Path) -> None:
        super().__init__()
        self.sources = sources
        self.destination = destination

    def run(self) -> None:
        """Convert, and hand back whatever came of it."""
        try:
            self.done.emit(convert(self.sources, self.destination))
        except Exception as exc:  # reported in the window, not on a dead console
            self.done.emit(exc)


class Window(QWidget):
    """The window itself, and the state it keeps: which files are queued."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"XML to XLSX {__version__}")
        self.setStyleSheet(stylesheet(PALETTE))
        self.resize(860, 620)
        self.setMinimumSize(600, 480)
        self.setAcceptDrops(True)

        self.sources: list[Path] = []
        # named with an underscore because QObject already has a thread()
        self._thread: QThread | None = None
        self._conversion: Conversion | None = None

        self._build()
        self._show_files()

    # building the window

    def _build(self) -> None:
        """Lay out the widgets, top to bottom."""
        heading = QLabel("Plate reader XML to Excel")
        heading.setObjectName("heading")

        self.zone = DropZone()
        self.zone.clicked.connect(self.choose_files)
        self.zone.dropped.connect(self.add_files)

        self.convert_button = QPushButton("Convert")
        self.convert_button.setObjectName("convert")
        self.convert_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_button.clicked.connect(self.start_convert)

        self.clear_button = QPushButton("Clear")
        self.clear_button.setObjectName("clear")
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.clicked.connect(self.clear_files)

        buttons = QHBoxLayout()
        buttons.addWidget(self.convert_button)
        buttons.addStretch()
        buttons.addWidget(self.clear_button)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # no total to count towards, so it just moves
        self.progress.setTextVisible(False)
        self.progress.hide()

        self.status = QLabel("")
        self.status.setObjectName("status")
        self.status.setWordWrap(True)

        self.table = self._build_table()
        self.table.hide()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 24, 26, 24)
        layout.setSpacing(16)
        layout.addWidget(heading)
        layout.addWidget(self.zone, stretch=2)
        layout.addLayout(buttons)
        layout.addWidget(self.progress)
        layout.addWidget(self.status)
        layout.addWidget(self.table, stretch=3)

    def _build_table(self) -> QTableWidget:
        """The report table: one row per plate, no editing, no selecting."""
        table = QTableWidget(0, len(COLUMNS))
        table.setHorizontalHeaderLabels([column.label for column in COLUMNS])
        table.verticalHeader().hide()
        table.setShowGrid(False)
        table.setAlternatingRowColors(False)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        header = table.horizontalHeader()
        header.setDefaultAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        for position, column in enumerate(COLUMNS):
            mode = (
                QHeaderView.ResizeMode.Stretch
                if column.stretch
                else QHeaderView.ResizeMode.ResizeToContents
            )
            header.setSectionResizeMode(position, mode)

        return table

    # choosing files

    def dragEnterEvent(self, a0: QDragEnterEvent) -> None:
        """Files may also be dropped anywhere on the window, not just the panel."""
        if a0.mimeData().hasUrls():
            a0.acceptProposedAction()

    def dropEvent(self, a0: QDropEvent) -> None:
        self.add_files(paths_from(a0))
        a0.acceptProposedAction()

    def choose_files(self) -> None:
        """Open the file dialog and queue whatever comes back."""
        chosen, _ = QFileDialog.getOpenFileNames(
            self, "Choose XML exports", "", "XML exports (*.xml);;All files (*)"
        )
        self.add_files([Path(one) for one in chosen])

    def add_files(self, paths: list[Path]) -> None:
        """
        Queue files, expanding folders and ignoring what is already there.
        """
        for path in paths:
            found = sorted(path.glob("*.xml")) if path.is_dir() else [path]
            self.sources.extend(one for one in found if one not in self.sources)

        self._show_files()

    def clear_files(self) -> None:
        """Forget the queued files and hide the last report."""
        self.sources = []
        self.status.clear()
        self.table.hide()
        self._show_files()

    def _show_files(self) -> None:
        """Put the current queue into the panel and enable Convert."""
        self.convert_button.setEnabled(bool(self.sources))

        if not self.sources:
            self.zone.show_text("Drop XML files here", "or click to choose them")
            return

        names = [path.name for path in self.sources[:NAMES_SHOWN]]
        if len(self.sources) > NAMES_SHOWN:
            names.append(f"and {len(self.sources) - NAMES_SHOWN} more")

        count = "1 file" if len(self.sources) == 1 else f"{len(self.sources)} files"
        self.zone.show_text(f"{count} ready", ", ".join(names))

    # converting

    def start_convert(self) -> None:
        """Ask where to save, then run the conversion on a worker thread."""
        if not self.sources:
            return

        first = self.sources[0]
        suggestion = (
            first.with_suffix(".xlsx") if len(self.sources) == 1 else "plates.xlsx"
        )
        chosen, _ = QFileDialog.getSaveFileName(
            self,
            "Save workbook as",
            str(first.parent / suggestion),
            "Excel workbook (*.xlsx)",
        )
        if not chosen:
            return

        destination = Path(chosen)
        self.convert_button.setEnabled(False)
        self.status.setText("Converting...")
        self.status.setStyleSheet(f"color: {PALETTE.text};")
        self.table.hide()
        self.progress.show()

        # both are kept on self, because a thread that goes out of scope
        # is collected mid-run
        self._thread = QThread(self)
        self._conversion = Conversion(list(self.sources), destination)
        self._conversion.moveToThread(self._thread)
        self._thread.started.connect(self._conversion.run)
        self._conversion.done.connect(lambda result: self._finish(result, destination))
        self._conversion.done.connect(self._thread.quit)
        self._thread.finished.connect(self._conversion.deleteLater)
        self._thread.start()

    def _finish(self, result: object, destination: Path) -> None:
        """Show what the worker produced, back on the window's own thread."""
        self.progress.hide()
        self.convert_button.setEnabled(True)

        if isinstance(result, Exception):
            self._show_error(result, destination)
        else:
            self._show_report(result, destination)  # ty: ignore[invalid-argument-type]

    def _show_error(self, error: Exception, destination: Path) -> None:
        """Explain a failed run in one line."""
        if isinstance(error, OSError):
            message = (
                f"Could not write {destination.name}. "
                "If it is open in Excel, close it and try again."
            )
        else:
            message = f"Something went wrong: {error}"

        self.status.setText(message)
        self.status.setStyleSheet(f"color: {Outcome.ERROR.colour};")

    def _show_report(self, report: list[tuple[str, ...]], destination: Path) -> None:
        """Show the summary line and one table row per plate."""
        summary = Summary.of(report)
        self.status.setText(summary.text(destination))
        colour = PALETTE.text if summary.written else Outcome.ERROR.colour
        self.status.setStyleSheet(f"color: {colour};")

        self.table.setRowCount(len(report))
        for position, row in enumerate(report):
            ink = QColor(Outcome.of(row[-1]).colour)
            for cell, column in enumerate(COLUMNS):
                item = QTableWidgetItem(row[column.index])
                item.setForeground(ink)
                item.setToolTip(row[column.index])
                self.table.setItem(position, cell, item)
        self.table.show()


def main() -> None:
    """Open the window."""
    app = QApplication(sys.argv)
    app.setApplicationName("xml2xlsx")
    app.setApplicationDisplayName("XML to XLSX")

    window = Window()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
