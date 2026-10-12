from __future__ import annotations

import json
import math
import os
import re
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

from PIL import Image, ImageChops, ImageDraw, ImageOps, ImageTk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from translations import LANGUAGES, translate


PALETTE = (
    "#ff4040",
    "#00b7ff",
    "#39c96b",
    "#ffb000",
    "#b45cff",
    "#ff65b5",
    "#00c9ad",
    "#f07828",
)
DPI_CHOICES = ("100", "200", "300", "400", "500", "600")
LASSO_PREVIEW_OPACITY = 15
VERTEX_MAGNET_RADIUS = 12
PROJECT_FORMAT = "image-analyzer-project"
PROJECT_VERSION = 1


@dataclass
class Region:
    name: str
    points: list[tuple[float, float]]
    color: str
    opacity: int = 40
    border_width: int = 2


def calculate_region_areas(
    image_size: tuple[int, int], regions: list[Region]
) -> tuple[list[int], int]:
    """Return raster pixel areas per region and for the union of all regions."""
    union = Image.new("L", image_size, 0)
    region_areas: list[int] = []

    for region in regions:
        mask = Image.new("L", image_size, 0)
        if len(region.points) >= 3:
            ImageDraw.Draw(mask).polygon(region.points, fill=255)
        region_areas.append(mask.histogram()[255])
        union = ImageChops.lighter(union, mask)

    union_area = union.histogram()[255]
    return region_areas, union_area


def calculate_area_percentage(area: int, base_area: int) -> float | None:
    if base_area <= 0:
        return None
    return area / base_area * 100


def make_annotated_image(image: Image.Image, regions: list[Region]) -> Image.Image:
    """Composite translucent region fills and opaque outlines over the source."""
    base = image.convert("RGBA")
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")

    for region in regions:
        if len(region.points) < 3:
            continue
        red, green, blue = tuple(
            int(region.color[index : index + 2], 16) for index in (1, 3, 5)
        )
        fill = (red, green, blue, round(255 * region.opacity / 100))
        draw.polygon(
            region.points,
            fill=fill,
            outline=(red, green, blue, 255),
            width=max(1, region.border_width),
        )

    return Image.alpha_composite(base, overlay).convert("RGB")


def make_region_export(
    image: Image.Image,
    regions: list[Region],
    selected_indices: list[int],
    include_intersections: bool = True,
    crop_to_regions: bool = True,
) -> Image.Image:
    """Render selected regions on the source image, optionally cropping the result."""
    selected_regions = [regions[index] for index in selected_indices]
    result = make_annotated_image(image, selected_regions)
    coverage = Image.new("L", image.size, 0)
    for region in selected_regions:
        mask = Image.new("L", image.size, 0)
        if len(region.points) >= 3:
            ImageDraw.Draw(mask).polygon(region.points, fill=1)
        coverage = ImageChops.add(coverage, mask)

    if not include_intersections:
        overlap_mask = coverage.point(lambda value: 255 if value > 1 else 0)
        result.paste(image.convert("RGB"), mask=overlap_mask)

    if crop_to_regions:
        bounds = coverage.point(lambda value: 255 if value else 0).getbbox()
        if bounds is not None:
            margin = max((region.border_width + 1) // 2 for region in selected_regions)
            left, top, right, bottom = bounds
            bounds = (
                max(0, left - margin),
                max(0, top - margin),
                min(image.width, right + margin),
                min(image.height, bottom + margin),
            )
            result = result.crop(bounds)
    return result


def save_project_file(
    project_path: Path,
    image_path: Path,
    image: Image.Image,
    regions: list[Region],
    base_region_index: int | None,
    dpi: int,
    language: str,
) -> None:
    try:
        image_reference = os.path.relpath(image_path.resolve(), project_path.parent.resolve())
    except ValueError:
        image_reference = str(image_path.resolve())
    project = {
        "format": PROJECT_FORMAT,
        "version": PROJECT_VERSION,
        "image": {
            "path": image_reference,
            "width": image.width,
            "height": image.height,
        },
        "regions": [
            {
                "name": region.name,
                "points": [[x, y] for x, y in region.points],
                "color": region.color,
                "opacity": region.opacity,
                "border_width": region.border_width,
            }
            for region in regions
        ],
        "base_region_index": base_region_index,
        "dpi": dpi,
        "language": language,
    }
    project_path.write_text(
        json.dumps(project, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def load_project_file(
    project_path: Path,
) -> tuple[Image.Image, Path, list[Region], int | None, int, str]:
    project = json.loads(project_path.read_text(encoding="utf-8"))
    if not isinstance(project, dict) or project.get("format") != PROJECT_FORMAT:
        raise ValueError("Not an Image Analyzer project file.")
    version = project.get("version")
    if not isinstance(version, int) or isinstance(version, bool) or version != PROJECT_VERSION:
        raise ValueError(f"Unsupported project version: {project.get('version')!r}")

    image_data = project.get("image")
    if not isinstance(image_data, dict) or not isinstance(image_data.get("path"), str):
        raise ValueError("Project does not contain a valid source image path.")
    image_width, image_height = image_data.get("width"), image_data.get("height")
    if (
        not isinstance(image_width, int)
        or isinstance(image_width, bool)
        or image_width <= 0
        or not isinstance(image_height, int)
        or isinstance(image_height, bool)
        or image_height <= 0
    ):
        raise ValueError("Project contains invalid source image dimensions.")
    image_path = Path(image_data["path"])
    if not image_path.is_absolute():
        image_path = project_path.parent / image_path
    image_path = image_path.resolve()
    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
    expected_size = (image_width, image_height)
    if expected_size != image.size:
        raise ValueError("The source image dimensions do not match the project.")

    region_data = project.get("regions")
    if not isinstance(region_data, list):
        raise ValueError("Project regions must be a list.")
    regions = []
    for item in region_data:
        if not isinstance(item, dict):
            raise ValueError("Project contains an invalid region.")
        name = item.get("name")
        color = item.get("color")
        points = item.get("points")
        opacity = item.get("opacity")
        border_width = item.get("border_width")
        if (
            not isinstance(name, str)
            or not name.strip()
            or not isinstance(color, str)
            or re.fullmatch(r"#[0-9a-fA-F]{6}", color) is None
            or not isinstance(points, list)
            or not isinstance(opacity, int)
            or isinstance(opacity, bool)
            or not 0 <= opacity <= 100
            or not isinstance(border_width, int)
            or isinstance(border_width, bool)
            or not 1 <= border_width <= 12
        ):
            raise ValueError("Project contains invalid region settings.")
        parsed_points: list[tuple[float, float]] = []
        for point in points:
            if (
                not isinstance(point, list)
                or len(point) != 2
                or any(
                    not isinstance(coordinate, (int, float))
                    or isinstance(coordinate, bool)
                    or (
                        isinstance(coordinate, float)
                        and not math.isfinite(coordinate)
                    )
                    for coordinate in point
                )
            ):
                raise ValueError("Project contains invalid region coordinates.")
            x, y = point
            if not 0 <= x < image.width or not 0 <= y < image.height:
                raise ValueError("Project contains coordinates outside the source image.")
            parsed_points.append((float(x), float(y)))
        if len(parsed_points) < 3:
            raise ValueError("Project regions must contain at least three points.")
        regions.append(
            Region(name, parsed_points, color.lower(), opacity, border_width)
        )

    base_region_index = project.get("base_region_index")
    if base_region_index is not None and (
        not isinstance(base_region_index, int)
        or isinstance(base_region_index, bool)
        or not 0 <= base_region_index < len(regions)
    ):
        raise ValueError("Project contains an invalid base region.")
    dpi = project.get("dpi", 300)
    if not isinstance(dpi, int) or dpi not in {100, 200, 300, 400, 500, 600}:
        raise ValueError("Project contains an invalid DPI value.")
    language = project.get("language", "ru")
    if not isinstance(language, str) or language not in {"ru", "sr", "en"}:
        raise ValueError("Project contains an unsupported interface language.")
    return image, image_path, regions, base_region_index, dpi, language


class ImageAnalyzerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.language = "ru"
        self._localized_widgets: list[tuple[tk.Widget, str]] = []
        self.root.title(self._t("title"))
        self.root.geometry("1280x820")
        self.root.minsize(850, 560)

        self.image: Image.Image | None = None
        self.image_path: Path | None = None
        self.project_path: Path | None = None
        self.tk_image: ImageTk.PhotoImage | None = None
        self.regions: list[Region] = []
        self.selected_index: int | None = None
        self.selected_vertex: int | None = None
        self.dragging_vertex: int | None = None
        self.hovered_vertex: tuple[int, int] | None = None
        self.suppress_vertex_hover = False
        self.delete_vertex_mode = False
        self.drawing = False
        self.lasso_points: list[tuple[float, float]] = []
        self.lasso_region_index: int | None = None
        self.scale = 1.0
        self.zoom = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.pan_start: tuple[float, float, float, float] | None = None
        self.image_x = 0.0
        self.image_y = 0.0
        self.scale_x = 1.0
        self.scale_y = 1.0
        self.display_size = (0, 0)
        self.region_areas: list[int] = []
        self.union_area = 0
        self.base_region_index: int | None = None
        self.base_area_var = tk.StringVar(value=self._t("whole_image"))
        self.base_area_options: dict[str, int | None] = {}
        self.language_var = tk.StringVar(value="Русский")
        self.opacity_var = tk.IntVar(value=60)
        self.border_width_var = tk.IntVar(value=2)
        self.opacity_value_var = tk.StringVar(value="60%")
        self.border_width_value_var = tk.StringVar(value="2 px")
        self.dpi_var = tk.StringVar(value="300")
        self.status_var = tk.StringVar(value=self._t("status_load_image"))

        self._build_ui()
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<Motion>", self._on_mouse_move)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<Button-4>", self._on_mouse_wheel_linux)
        self.canvas.bind("<Button-5>", self._on_mouse_wheel_linux)
        self.canvas.bind("<ButtonPress-3>", self._on_pan_start)
        self.canvas.bind("<B3-Motion>", self._on_pan_drag)
        self.canvas.bind("<ButtonRelease-3>", self._on_pan_end)
        self.root.bind("<Delete>", self._delete_selected_region)
        self.root.bind("<Escape>", self._on_escape)

    def _build_ui(self) -> None:
        toolbar = ttk.Frame(self.root, padding=(8, 8, 8, 4))
        toolbar.pack(side=tk.TOP, fill=tk.X)
        self._button(toolbar, "load_image", command=self.load_image).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self._button(toolbar, "open_project", command=self.open_project).pack(
            side=tk.LEFT, padx=6
        )
        self._button(toolbar, "save_project", command=self.save_project).pack(
            side=tk.LEFT, padx=6
        )
        self._button(toolbar, "new_lasso", command=self.start_lasso).pack(
            side=tk.LEFT, padx=6
        )
        self.delete_region_button = self._button(
            toolbar, "delete_region", command=self.delete_selected_region
        )
        self.delete_region_button.pack(side=tk.LEFT, padx=6)
        self._button(toolbar, "export_image", command=self.export_image).pack(
            side=tk.LEFT, padx=6
        )
        self._button(toolbar, "export_excel", command=self.export_excel).pack(
            side=tk.LEFT, padx=6
        )
        self._label(toolbar, "dpi").pack(side=tk.LEFT, padx=(16, 4))
        ttk.Combobox(
            toolbar,
            textvariable=self.dpi_var,
            values=DPI_CHOICES,
            state="readonly",
            width=5,
        ).pack(side=tk.LEFT)
        self._label(toolbar, "language").pack(side=tk.RIGHT, padx=(8, 4))
        self.language_combo = ttk.Combobox(
            toolbar,
            textvariable=self.language_var,
            values=tuple(LANGUAGES),
            state="readonly",
            width=10,
        )
        self.language_combo.pack(side=tk.RIGHT)
        self.language_combo.bind("<<ComboboxSelected>>", self._on_language_change)

        body = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        side = ttk.Frame(body, width=320, padding=8)
        body.add(side, weight=0)
        self._label(side, "regions", font=("", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 5)
        )
        self._label(side, "base_area_prompt").pack(anchor=tk.W)
        self.base_area_combo = ttk.Combobox(
            side,
            textvariable=self.base_area_var,
            state="readonly",
        )
        self.base_area_combo.pack(fill=tk.X, pady=(3, 6))
        self.base_area_combo.bind("<<ComboboxSelected>>", self._on_base_area_change)
        self.region_list = tk.Listbox(side, height=13, exportselection=False)
        self.region_list.pack(fill=tk.BOTH, expand=True)
        self.region_list.bind("<Button-1>", self._on_region_click)
        self.region_list.bind("<<ListboxSelect>>", self._on_region_select)
        self.region_list.bind("<Escape>", self._on_escape)

        region_buttons = ttk.Frame(side)
        region_buttons.pack(fill=tk.X, pady=(6, 10))
        self.rename_region_button = self._button(
            region_buttons, "rename", command=self.rename_region
        )
        self.rename_region_button.pack(side=tk.LEFT, padx=(0, 5))
        self.choose_color_button = self._button(
            region_buttons, "color", command=self.choose_color
        )
        self.choose_color_button.pack(side=tk.LEFT, padx=5)
        self.delete_vertex_button = self._button(
            region_buttons, "delete_vertex", command=self.toggle_delete_vertex_mode
        )
        self.delete_vertex_button.pack(side=tk.LEFT, padx=5)

        self._label(side, "fill_transparency").pack(anchor=tk.W)
        opacity_control_row = ttk.Frame(side)
        opacity_control_row.pack(fill=tk.X, pady=(3, 10))
        self.opacity_scale = ttk.Scale(
            opacity_control_row,
            from_=10,
            to=95,
            variable=self.opacity_var,
            command=self._on_opacity_change,
        )
        self.opacity_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            opacity_control_row,
            textvariable=self.opacity_value_var,
            width=5,
            anchor=tk.E,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self._label(side, "border_width").pack(anchor=tk.W)
        border_width_control_row = ttk.Frame(side)
        border_width_control_row.pack(fill=tk.X, pady=(3, 10))
        self.border_width_scale = ttk.Scale(
            border_width_control_row,
            from_=1,
            to=12,
            variable=self.border_width_var,
            command=self._on_border_width_change,
        )
        self.border_width_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            border_width_control_row,
            textvariable=self.border_width_value_var,
            width=6,
            anchor=tk.E,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self._sync_region_controls()

        ttk.Separator(side).pack(fill=tk.X, pady=6)
        self._label(side, "area_summary", font=("", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 5)
        )
        self.summary_label = ttk.Label(
            side, text=self._t("summary_empty")
        )
        self.summary_label.pack(anchor=tk.W)
        self.stats_label = ttk.Label(side, text="", justify=tk.LEFT)
        self.stats_label.pack(anchor=tk.W, pady=(8, 0))

        canvas_frame = ttk.Frame(body)
        body.add(canvas_frame, weight=1)
        self.canvas = tk.Canvas(canvas_frame, background="#252525", highlightthickness=0)
        self.horizontal_scrollbar = ttk.Scrollbar(
            canvas_frame,
            orient=tk.HORIZONTAL,
            command=lambda *args: self._scroll_view("x", *args),
        )
        self.vertical_scrollbar = ttk.Scrollbar(
            canvas_frame,
            orient=tk.VERTICAL,
            command=lambda *args: self._scroll_view("y", *args),
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        self.horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        ttk.Frame(canvas_frame).grid(row=1, column=1)
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.root,
            textvariable=self.status_var,
            anchor=tk.W,
            relief=tk.SUNKEN,
            padding=(8, 4),
        ).pack(side=tk.BOTTOM, fill=tk.X)

    def _t(self, key: str, **values: object) -> str:
        return translate(getattr(self, "language", "ru"), key, **values)

    def _label(self, parent: tk.Misc, key: str, **options: object) -> ttk.Label:
        widget = ttk.Label(parent, text=self._t(key), **options)
        self._localized_widgets.append((widget, key))
        return widget

    def _button(self, parent: tk.Misc, key: str, **options: object) -> ttk.Button:
        widget = ttk.Button(parent, text=self._t(key), **options)
        self._localized_widgets.append((widget, key))
        return widget

    def _on_language_change(self, _event: tk.Event | None = None) -> None:
        self.language = LANGUAGES[self.language_var.get()]
        self.apply_language()

    def apply_language(self) -> None:
        self.root.title(self._t("title"))
        for widget, key in self._localized_widgets:
            widget.configure(text=self._t(key))
        self._update_delete_vertex_button()
        self.status_var.set(self._t("language_changed"))
        self._refresh_regions()
        self._update_statistics()
        self._render()

    def _update_delete_vertex_button(self) -> None:
        if hasattr(self, "delete_vertex_button"):
            key = "cancel_delete" if self.delete_vertex_mode else "delete_vertex"
            self.delete_vertex_button.configure(text=self._t(key))

    def load_image(self) -> None:
        path = filedialog.askopenfilename(
            title=self._t("select_image"),
            filetypes=[
                (self._t("image_files"), "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                (self._t("all_files"), "*.*"),
            ],
        )
        if not path:
            return
        try:
            with Image.open(path) as source:
                loaded = ImageOps.exif_transpose(source).convert("RGB")
        except (OSError, ValueError) as exc:
            messagebox.showerror(self._t("open_image_error"), str(exc))
            return

        self.image = loaded
        self.image_path = Path(path)
        self.project_path = None
        self.regions.clear()
        self.delete_vertex_mode = False
        self._update_delete_vertex_button()
        self.selected_index = None
        self.selected_vertex = None
        self._sync_region_controls()
        self.base_region_index = None
        self.drawing = False
        self.lasso_points.clear()
        self.lasso_region_index = None
        self.hovered_vertex = None
        self.zoom = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.pan_start = None
        self.canvas.configure(cursor="arrow")
        self._refresh_regions()
        self._update_statistics()
        self._render()
        self.status_var.set(
            self._t(
                "image_loaded",
                name=self.image_path.name,
                width=loaded.width,
                height=loaded.height,
            )
        )

    def save_project(self) -> None:
        if self.image is None or self.image_path is None:
            messagebox.showinfo(self._t("no_image"), self._t("load_first"))
            return
        path = filedialog.asksaveasfilename(
            title=self._t("save_project_title"),
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        project_path = Path(path)
        try:
            save_project_file(
                project_path,
                self.image_path,
                self.image,
                self.regions,
                self.base_region_index,
                int(self.dpi_var.get()),
                self.language,
            )
        except (OSError, ValueError, TypeError) as exc:
            messagebox.showerror(self._t("project_save_error"), str(exc))
            return
        self.project_path = project_path
        self.status_var.set(
            self._t("project_saved", name=project_path.name)
        )

    def open_project(self) -> None:
        path = filedialog.askopenfilename(
            title=self._t("open_project_title"),
            filetypes=[("JSON", "*.json"), (self._t("all_files"), "*.*")],
        )
        if not path:
            return
        project_path = Path(path)
        try:
            (
                image,
                image_path,
                regions,
                base_region_index,
                dpi,
                language,
            ) = load_project_file(project_path)
        except OSError as exc:
            messagebox.showerror(self._t("project_open_error"), str(exc))
            return
        except (ValueError, TypeError, KeyError):
            messagebox.showerror(
                self._t("project_open_error"), self._t("project_invalid")
            )
            return

        self.image = image
        self.image_path = image_path
        self.project_path = project_path
        self.regions = regions
        self.base_region_index = base_region_index
        self.dpi_var.set(str(dpi))
        self.language = language
        self.language_var.set(next(name for name, code in LANGUAGES.items() if code == language))
        self.delete_vertex_mode = False
        self._update_delete_vertex_button()
        self.selected_index = None
        self.selected_vertex = None
        self.dragging_vertex = None
        self.hovered_vertex = None
        self.drawing = False
        self.lasso_points.clear()
        self.lasso_region_index = None
        self.zoom = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.pan_start = None
        self.canvas.configure(cursor="arrow")
        self._sync_region_controls()
        self._update_statistics()
        self.apply_language()
        self.status_var.set(
            self._t(
                "project_opened",
                name=project_path.name,
                image=image_path.name,
            )
        )

    def start_lasso(self) -> None:
        if self.image is None:
            messagebox.showinfo(self._t("no_image"), self._t("load_first"))
            return
        self._set_delete_vertex_mode(False)
        if self.drawing:
            self._finish_lasso()
        self.region_list.selection_clear(0, tk.END)
        self.selected_index = None
        self._sync_region_controls()
        self.suppress_vertex_hover = False
        self.drawing = True
        self.lasso_points.clear()
        self.lasso_region_index = None
        self.hovered_vertex = None
        self.selected_vertex = None
        self.dragging_vertex = None
        self.canvas.configure(cursor="crosshair")
        self.status_var.set(self._t("lasso_instructions"))
        self._render()

    def _on_canvas_resize(self, _event: tk.Event) -> None:
        if self.image is not None:
            self._render()

    def _fit_image(self) -> None:
        assert self.image is not None
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        fit_scale = min(canvas_width / self.image.width, canvas_height / self.image.height)
        self.scale = fit_scale * self.zoom
        display_width = max(1, round(self.image.width * self.scale))
        display_height = max(1, round(self.image.height * self.scale))
        self.display_size = (display_width, display_height)
        self.scale_x = display_width / self.image.width
        self.scale_y = display_height / self.image.height
        self.image_x = (canvas_width - display_width) / 2 + self.pan_x
        self.image_y = (canvas_height - display_height) / 2 + self.pan_y

    def _render(self) -> None:
        self.canvas.delete("all")
        if self.image is None:
            if hasattr(self, "horizontal_scrollbar"):
                self.horizontal_scrollbar.set(0, 1)
                self.vertical_scrollbar.set(0, 1)
            self.canvas.create_text(
                max(1, self.canvas.winfo_width()) / 2,
                max(1, self.canvas.winfo_height()) / 2,
                text=self._t("canvas_empty"),
                fill="white",
                font=("", 16),
            )
            return

        self._fit_image()
        self._update_scrollbars()
        preview_source = self.image.resize(self.display_size, Image.Resampling.LANCZOS)
        scale_x = self.display_size[0] / self.image.width
        scale_y = self.display_size[1] / self.image.height
        preview_regions = [
            Region(
                region.name,
                [(x * scale_x, y * scale_y) for x, y in region.points],
                region.color,
                region.opacity,
                max(1, round(region.border_width * (scale_x + scale_y) / 2)),
            )
            for region in self.regions
        ]
        preview = make_annotated_image(preview_source, preview_regions)
        self.tk_image = ImageTk.PhotoImage(preview)
        self.canvas.create_image(self.image_x, self.image_y, image=self.tk_image, anchor=tk.NW)

        if self.drawing and self.lasso_points:
            display_points = [self._image_to_canvas(point) for point in self.lasso_points]
            if len(display_points) > 1:
                if len(display_points) >= 3:
                    display_points.append(display_points[0])
                self.canvas.create_line(
                    *[coord for point in display_points for coord in point],
                    fill="#ffffff",
                    width=2,
                    dash=(5, 3),
                )
            for x, y in display_points:
                self.canvas.create_oval(
                    x - 4, y - 4, x + 4, y + 4, fill="#ffe600", outline="#101010", width=1
                )

        if self.selected_index is not None and self.selected_index < len(self.regions):
            region = self.regions[self.selected_index]
            for index, point in enumerate(region.points):
                x, y = self._image_to_canvas(point)
                is_hovered = self.hovered_vertex == (self.selected_index, index)
                is_selected = index == self.selected_vertex
                radius = 7 if is_hovered or is_selected else 5
                self.canvas.create_oval(
                    x - radius,
                    y - radius,
                    x + radius,
                    y + radius,
                    fill="#ffffff" if is_hovered or is_selected else region.color,
                    outline="#101010",
                    width=2,
                )

    def _update_scrollbars(self) -> None:
        if not hasattr(self, "horizontal_scrollbar"):
            return
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        display_width, display_height = self.display_size

        if display_width <= canvas_width:
            self.horizontal_scrollbar.set(0, 1)
        else:
            first = max(0.0, min(1.0, -self.image_x / display_width))
            last = min(1.0, first + canvas_width / display_width)
            self.horizontal_scrollbar.set(first, last)

        if display_height <= canvas_height:
            self.vertical_scrollbar.set(0, 1)
        else:
            first = max(0.0, min(1.0, -self.image_y / display_height))
            last = min(1.0, first + canvas_height / display_height)
            self.vertical_scrollbar.set(first, last)

    def _scroll_view(self, axis: str, command: str, value: str, *extra: str) -> None:
        if self.image is None or axis not in {"x", "y"}:
            return
        horizontal = axis == "x"
        display_size = self.display_size[0 if horizontal else 1]
        viewport_size = max(
            1,
            self.canvas.winfo_width() if horizontal else self.canvas.winfo_height(),
        )
        if display_size <= viewport_size:
            return

        if command == "moveto":
            first = max(0.0, min(1.0, float(value)))
            first = min(first, 1.0 - viewport_size / display_size)
            origin = -first * display_size
            centered_origin = (viewport_size - display_size) / 2
            pan = origin - centered_origin
        elif command == "scroll":
            amount = float(value)
            step = viewport_size * (0.9 if extra and extra[0] == "pages" else 0.1)
            pan = (self.pan_x if horizontal else self.pan_y) - amount * step
        else:
            return

        if horizontal:
            self.pan_x = pan
        else:
            self.pan_y = pan
        self._render()

    def _image_to_canvas(self, point: tuple[float, float]) -> tuple[float, float]:
        return self.image_x + point[0] * self.scale_x, self.image_y + point[1] * self.scale_y

    def _canvas_to_image(self, x: float, y: float) -> tuple[float, float] | None:
        if self.image is None:
            return None
        image_x = (x - self.image_x) / self.scale_x
        image_y = (y - self.image_y) / self.scale_y
        if not (0 <= image_x < self.image.width and 0 <= image_y < self.image.height):
            return None
        return image_x, image_y

    def _on_mouse_wheel(self, event: tk.Event) -> None:
        if self.image is None or event.delta == 0:
            return
        self._zoom_at(event.x, event.y, 1.15 ** (event.delta / 120))

    def _on_mouse_wheel_linux(self, event: tk.Event) -> None:
        if event.num == 4:
            self._zoom_at(event.x, event.y, 1.15)
        elif event.num == 5:
            self._zoom_at(event.x, event.y, 1 / 1.15)

    def _zoom_at(self, x: float, y: float, factor: float) -> None:
        if self.image is None:
            return
        image_x = (x - self.image_x) / self.scale_x
        image_y = (y - self.image_y) / self.scale_y
        self.zoom = max(0.1, min(20.0, self.zoom * factor))
        self._fit_image()
        self.pan_x += x - (self.image_x + image_x * self.scale_x)
        self.pan_y += y - (self.image_y + image_y * self.scale_y)
        self._render()

    def _on_pan_start(self, event: tk.Event) -> None:
        if self.image is None:
            return
        self.pan_start = (event.x, event.y, self.pan_x, self.pan_y)
        self._set_canvas_cursor()

    def _on_pan_drag(self, event: tk.Event) -> None:
        if self.pan_start is None:
            return
        start_x, start_y, start_pan_x, start_pan_y = self.pan_start
        self.pan_x = start_pan_x + event.x - start_x
        self.pan_y = start_pan_y + event.y - start_y
        self._render()

    def _on_pan_end(self, _event: tk.Event) -> None:
        self.pan_start = None
        self._set_canvas_cursor()

    def _on_press(self, event: tk.Event) -> None:
        if self.delete_vertex_mode:
            nearest = (
                self._find_nearby_vertex(event.x, event.y, self.selected_index)
                if self.selected_index is not None
                else None
            )
            if nearest is None:
                self.hovered_vertex = None
                self.selected_vertex = None
                self.status_var.set(
                    self._t("click_vertex_to_delete")
                )
                self._render()
                return
            self.hovered_vertex = nearest
            self.selected_index, self.selected_vertex = nearest
            self.region_list.selection_clear(0, tk.END)
            self.region_list.selection_set(self.selected_index)
            self.region_list.activate(self.selected_index)
            point_count = len(self.regions[self.selected_index].points)
            self.delete_selected_vertex()
            if self.delete_vertex_mode and len(
                self.regions[self.selected_index].points
            ) < point_count:
                self.status_var.set(
                    self._t("vertex_deleted_next")
                )
            return
        if self.drawing:
            nearest = (
                self._find_nearby_vertex(event.x, event.y, self.lasso_region_index)
                if self.lasso_region_index is not None
                else None
            )
            if nearest is not None:
                self.selected_index, self.selected_vertex = nearest
                self.hovered_vertex = nearest
                self.dragging_vertex = nearest[1]
                self._sync_border_width_control()
                self._sync_region_controls()
                self._set_canvas_cursor()
                self._render()
                return
            point = self._canvas_to_image(event.x, event.y)
            if point is not None:
                if self.lasso_points:
                    last_x, last_y = self._image_to_canvas(self.lasso_points[-1])
                    if math.hypot(event.x - last_x, event.y - last_y) <= 3:
                        return
                self._add_lasso_point(point)
                self._render()
            return
        self.dragging_vertex = None
        nearest = (
            self._find_nearby_vertex(event.x, event.y, self.selected_index)
            if self.selected_index is not None
            else None
        )
        self.hovered_vertex = nearest
        if nearest is not None:
            self.selected_vertex = nearest[1]
            self.dragging_vertex = self.selected_vertex
        elif self.selected_index is not None:
            point = self._canvas_to_image(event.x, event.y)
            if point is not None and self._insert_vertex_near_edge(point):
                return
            self.selected_vertex = None
        else:
            self.selected_vertex = None
        self._render()

    def _insert_vertex_near_edge(self, point: tuple[float, float]) -> bool:
        if self.selected_index is None:
            return False
        region = self.regions[self.selected_index]
        if len(region.points) < 3:
            return False

        closest_edge: tuple[float, int, tuple[float, float]] | None = None
        for index, start in enumerate(region.points):
            end = region.points[(index + 1) % len(region.points)]
            dx, dy = end[0] - start[0], end[1] - start[1]
            length_squared = dx * dx + dy * dy
            t = 0 if length_squared == 0 else max(
                0,
                min(
                    1,
                    ((point[0] - start[0]) * dx + (point[1] - start[1]) * dy)
                    / length_squared,
                ),
            )
            projected = (start[0] + t * dx, start[1] + t * dy)
            distance = math.hypot(
                (point[0] - projected[0]) * self.scale_x,
                (point[1] - projected[1]) * self.scale_y,
            )
            if closest_edge is None or distance < closest_edge[0]:
                closest_edge = (distance, index + 1, projected)

        if closest_edge is None or closest_edge[0] > VERTEX_MAGNET_RADIUS:
            return False
        _, insert_at, projected = closest_edge
        region.points.insert(insert_at, projected)
        self.selected_vertex = insert_at
        self.hovered_vertex = (self.selected_index, insert_at)
        self._update_statistics()
        self._refresh_regions()
        self._render()
        return True

    def _on_drag(self, event: tk.Event) -> None:
        point = self._canvas_to_image(event.x, event.y)
        if (
            self.dragging_vertex is not None
            and point is not None
            and self.selected_index is not None
        ):
            self.regions[self.selected_index].points[self.dragging_vertex] = point
            if self.drawing and self.selected_index == self.lasso_region_index:
                self.lasso_points = self.regions[self.selected_index].points
            self._render()
            return
        if self.drawing:
            if point is not None and self.lasso_points:
                last = self._image_to_canvas(self.lasso_points[-1])
                if math.hypot(event.x - last[0], event.y - last[1]) >= 3:
                    self._add_lasso_point(point)
                    self._render()
            return

    def _on_release(self, _event: tk.Event) -> None:
        if self.dragging_vertex is not None:
            self.dragging_vertex = None
            self._update_statistics()
            self._refresh_regions()
            self.hovered_vertex = self._find_nearby_vertex(
                _event.x,
                _event.y,
                (
                    self.lasso_region_index
                    if self.drawing
                    else self.selected_index
                ),
            )
            if self.hovered_vertex is not None:
                self.selected_vertex = self.hovered_vertex[1]
            else:
                self.selected_vertex = None
            self._set_canvas_cursor()
            self._render()
        elif self.drawing:
            return

    def _on_double_click(self, event: tk.Event) -> None:
        if self.drawing:
            if self.dragging_vertex is not None:
                self.dragging_vertex = None
                self._finish_lasso()
                return
            point = self._canvas_to_image(event.x, event.y)
            if point is not None and self.lasso_points:
                last = self._image_to_canvas(self.lasso_points[-1])
                if math.hypot(event.x - last[0], event.y - last[1]) > 3:
                    self._add_lasso_point(point)
            self._finish_lasso()
            return
        if self.selected_index is None:
            return
        region = self.regions[self.selected_index]
        point = self._canvas_to_image(event.x, event.y)
        if point is None or len(region.points) < 3:
            return
        if self._find_nearby_vertex(
            event.x, event.y, region_index=self.selected_index
        ) is not None:
            return
        if self._insert_vertex_near_edge(point):
            return

    def _on_mouse_move(self, event: tk.Event) -> None:
        if self.image is None or self.dragging_vertex is not None:
            return
        if getattr(self, "suppress_vertex_hover", False):
            nearby_after_escape = self._find_nearby_vertex(event.x, event.y)
            if nearby_after_escape is not None:
                return
            self.suppress_vertex_hover = False
        active_region_index = (
            self.lasso_region_index if self.drawing else self.selected_index
        )
        nearby = (
            self._find_nearby_vertex(
                event.x, event.y, active_region_index
            )
            if active_region_index is not None
            else None
        )
        if nearby == self.hovered_vertex:
            return
        self.hovered_vertex = nearby
        if nearby is not None:
            self.selected_index, self.selected_vertex = nearby
            self.region_list.selection_clear(0, tk.END)
            self.region_list.selection_set(self.selected_index)
            self.region_list.activate(self.selected_index)
            self._sync_opacity_control()
            self._sync_border_width_control()
            if self.delete_vertex_mode:
                self.status_var.set(
                    self._t("vertex_click_delete")
                )
            else:
                self.status_var.set(
                    self._t("vertex_drag")
                )
        else:
            self.selected_vertex = None
            if self.delete_vertex_mode:
                self.status_var.set(
                    self._t("delete_mode_hover")
                )
            else:
                self.status_var.set(self._t("hover_vertex_edit"))
            self._sync_region_controls()
        self._set_canvas_cursor()
        self._render()

    def _find_nearby_vertex(
        self, x: float, y: float, region_index: int | None = None
    ) -> tuple[int, int] | None:
        closest: tuple[float, int, int] | None = None
        for current_region_index, region in enumerate(self.regions):
            if region_index is not None and current_region_index != region_index:
                continue
            for vertex_index, point in enumerate(region.points):
                canvas_x, canvas_y = self._image_to_canvas(point)
                distance = math.hypot(x - canvas_x, y - canvas_y)
                if distance <= VERTEX_MAGNET_RADIUS and (
                    closest is None or distance < closest[0]
                ):
                    closest = (distance, current_region_index, vertex_index)
        if closest is None:
            return None
        return closest[1], closest[2]

    def _set_canvas_cursor(self) -> None:
        if self.pan_start is not None:
            cursor = "fleur"
        elif self.delete_vertex_mode:
            cursor = "crosshair"
        elif self.dragging_vertex is not None or self.hovered_vertex is not None:
            cursor = "fleur"
        elif self.drawing:
            cursor = "crosshair"
        else:
            cursor = "arrow"
        self.canvas.configure(cursor=cursor)

    def toggle_delete_vertex_mode(self) -> None:
        self._set_delete_vertex_mode(not self.delete_vertex_mode)

    def _set_delete_vertex_mode(self, enabled: bool) -> None:
        self.delete_vertex_mode = enabled
        if hasattr(self, "delete_vertex_button"):
            self.delete_vertex_button.configure(
                text=self._t("cancel_delete" if enabled else "delete_vertex")
            )
        if enabled:
            if self.drawing:
                self._finish_lasso()
            self.hovered_vertex = None
            self.selected_vertex = None
            self.status_var.set(
                self._t("delete_mode_hint")
            )
        else:
            self.hovered_vertex = None
            self.selected_vertex = None
            if hasattr(self, "status_var"):
                self.status_var.set(self._t("delete_mode_off"))
        if hasattr(self, "canvas"):
            self._set_canvas_cursor()
            self._render()

    def _add_lasso_point(self, point: tuple[float, float]) -> None:
        self.lasso_points.append(point)
        if len(self.lasso_points) == 3:
            number = len(self.regions) + 1
            existing_names = {region.name for region in self.regions}
            while self._t("region_name", number=number) in existing_names:
                number += 1
            region = Region(
                name=self._t("region_name", number=number),
                points=self.lasso_points,
                color=PALETTE[len(self.regions) % len(PALETTE)],
                opacity=LASSO_PREVIEW_OPACITY,
            )
            self.regions.append(region)
            self.lasso_region_index = len(self.regions) - 1
            self.selected_index = self.lasso_region_index
            self.selected_vertex = None
            self._sync_opacity_control()
            self._sync_border_width_control()
            self._sync_region_controls()
            self.region_list.selection_clear(0, tk.END)
            self._update_statistics()
            self._refresh_regions()
        elif self.lasso_region_index is not None:
            self._update_statistics()
            self._refresh_regions()

    def _finish_lasso(self) -> None:
        self.drawing = False
        self.lasso_points = []
        if self.lasso_region_index is not None:
            region = self.regions[self.lasso_region_index]
            self.selected_index = self.lasso_region_index
            self.selected_vertex = None
            self._sync_opacity_control()
            self._sync_border_width_control()
            self._sync_region_controls()
            self._update_statistics()
            self._refresh_regions()
            self.status_var.set(self._t("added_region", name=region.name))
        else:
            self.status_var.set(self._t("region_not_created"))
        self.lasso_region_index = None
        self._set_canvas_cursor()
        self._render()

    def _on_region_select(self, _event: tk.Event | None = None) -> None:
        selected = self.region_list.curselection()
        self.selected_index = selected[0] if selected else None
        if (
            self.hovered_vertex is not None
            and self.hovered_vertex[0] != self.selected_index
        ):
            self.hovered_vertex = None
        if self.hovered_vertex is not None and self.hovered_vertex[0] == self.selected_index:
            self.selected_vertex = self.hovered_vertex[1]
        else:
            self.selected_vertex = None
        self.dragging_vertex = None
        if self.selected_index is not None:
            self._sync_opacity_control()
            self._sync_border_width_control()
        self._sync_region_controls()
        self._render()

    def _on_escape(self, _event: tk.Event | None = None) -> str:
        if self.delete_vertex_mode:
            self._set_delete_vertex_mode(False)
        self._clear_region_selection(suppress_hover=True)
        return "break"

    def _clear_region_selection(self, *, suppress_hover: bool = False) -> None:
        if self.drawing:
            self._finish_lasso()
        self.drawing = False
        self.lasso_points = []
        self.lasso_region_index = None
        self.region_list.selection_clear(0, tk.END)
        self.selected_index = None
        self.selected_vertex = None
        self.hovered_vertex = None
        self.dragging_vertex = None
        self.suppress_vertex_hover = suppress_hover
        self._sync_region_controls()
        self._set_canvas_cursor()
        self._render()

    def _on_region_click(self, event: tk.Event) -> str:
        index = self.region_list.index(f"@{event.x},{event.y}")
        bounds = self.region_list.bbox(index) if 0 <= index < len(self.regions) else None
        if (
            bounds is None
            or event.y < bounds[1]
            or event.y >= bounds[1] + bounds[3]
            or not (0 <= index < len(self.regions))
        ):
            if self.delete_vertex_mode:
                self._set_delete_vertex_mode(False)
            self._clear_region_selection()
            return "break"

        if index == self.selected_index:
            if self.delete_vertex_mode:
                self._set_delete_vertex_mode(False)
            self._clear_region_selection()
            return "break"

        self.region_list.selection_clear(0, tk.END)
        self.region_list.selection_set(index)
        self.region_list.activate(index)
        self.hovered_vertex = None
        self._on_region_select()
        return "break"

    def _refresh_regions(self) -> None:
        self.region_list.delete(0, tk.END)
        self.base_area_options = {self._t("whole_image"): None}
        for index, region in enumerate(self.regions):
            label = self._t(
                "base_region_option", number=index + 1, name=region.name
            )
            self.base_area_options[label] = index
        self.base_area_combo.configure(values=tuple(self.base_area_options))
        if self.base_region_index is not None and self.base_region_index >= len(self.regions):
            self.base_region_index = None
        selected_base = next(
            (
                label
                for label, index in self.base_area_options.items()
                if index == self.base_region_index
            ),
            self._t("whole_image"),
        )
        self.base_area_var.set(selected_base)
        base_area = self._base_area_pixels()
        for index, region in enumerate(self.regions):
            area = self.region_areas[index] if index < len(self.region_areas) else 0
            percent = calculate_area_percentage(area, base_area)
            percent_text = "—" if percent is None else f"{percent:.2f}%"
            self.region_list.insert(tk.END, f"{region.name} — {percent_text}")
            self.region_list.itemconfigure(index, foreground=region.color)
        if self.selected_index is not None and self.selected_index < len(self.regions):
            self.region_list.selection_set(self.selected_index)
            self.region_list.activate(self.selected_index)

    def _update_statistics(self) -> None:
        if self.image is None:
            self.region_areas = []
            self.union_area = 0
            if hasattr(self, "summary_label"):
                self.summary_label.configure(text=self._t("summary_empty"))
            if hasattr(self, "stats_label"):
                self.stats_label.configure(text="")
            return
        self.region_areas, self.union_area = calculate_region_areas(
            self.image.size, self.regions
        )
        base_area = self._base_area_pixels()
        total_percent = calculate_area_percentage(sum(self.region_areas), base_area)
        union_percent = calculate_area_percentage(self.union_area, base_area)
        total_text = "—" if total_percent is None else f"{total_percent:.2f}%"
        union_text = "—" if union_percent is None else f"{union_percent:.2f}%"
        self.summary_label.configure(
            text="\n".join(
                (
                    self._t(
                        "summary_base", name=self._base_area_label(), area=base_area
                    ),
                    self._t("summary_sum", value=total_text),
                    self._t("summary_union", value=union_text),
                )
            )
        )
        self.stats_label.configure(
            text="\n".join(
                self._t(
                    "region_stat",
                    name=region.name,
                    percentage=self._format_percentage(area, base_area),
                    area=area,
                )
                for region, area in zip(self.regions, self.region_areas)
            )
        )

    @staticmethod
    def _format_percentage(area: int, base_area: int) -> str:
        percentage = calculate_area_percentage(area, base_area)
        return "—" if percentage is None else f"{percentage:.2f}%"

    def _base_area_pixels(self) -> int:
        if self.image is None:
            return 0
        if self.base_region_index is None:
            return self.image.width * self.image.height
        if self.base_region_index >= len(self.region_areas):
            return 0
        return self.region_areas[self.base_region_index]

    def _base_area_label(self) -> str:
        if self.base_region_index is None:
            return self._t("whole_image")
        if self.base_region_index >= len(self.regions):
            return self._t("whole_image")
        return self.regions[self.base_region_index].name

    def _on_base_area_change(self, _event: tk.Event | None = None) -> None:
        self.base_region_index = self.base_area_options.get(self.base_area_var.get())
        self._update_statistics()
        self._refresh_regions()

    def _on_opacity_change(self, _value: str) -> None:
        if self.selected_index is None or not (0 <= self.selected_index < len(self.regions)):
            return
        transparency = max(10, min(95, round(float(_value))))
        self.regions[self.selected_index].opacity = 100 - transparency
        if hasattr(self, "opacity_value_var"):
            self.opacity_value_var.set(f"{transparency}%")
        self._render()

    def _sync_opacity_control(self) -> None:
        if self.selected_index is None or not (0 <= self.selected_index < len(self.regions)):
            return
        self.opacity_var.set(100 - self.regions[self.selected_index].opacity)
        if hasattr(self, "opacity_value_var"):
            self.opacity_value_var.set(
                f"{100 - self.regions[self.selected_index].opacity}%"
            )

    def _on_border_width_change(self, value: str) -> None:
        if self.selected_index is None or not (0 <= self.selected_index < len(self.regions)):
            return
        self.regions[self.selected_index].border_width = max(
            1, min(12, round(float(value)))
        )
        if hasattr(self, "border_width_value_var"):
            self.border_width_value_var.set(
                f"{self.regions[self.selected_index].border_width} px"
            )
        self._render()

    def _sync_border_width_control(self) -> None:
        if self.selected_index is None or not (0 <= self.selected_index < len(self.regions)):
            return
        self.border_width_var.set(self.regions[self.selected_index].border_width)
        if hasattr(self, "border_width_value_var"):
            self.border_width_value_var.set(
                f"{self.regions[self.selected_index].border_width} px"
            )

    def _sync_region_controls(self) -> None:
        enabled = self.selected_index is not None and 0 <= self.selected_index < len(
            self.regions
        )
        state = tk.NORMAL if enabled else tk.DISABLED
        for widget_name in (
            "delete_region_button",
            "rename_region_button",
            "choose_color_button",
            "delete_vertex_button",
            "opacity_scale",
            "border_width_scale",
        ):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                widget.configure(state=state)
        if hasattr(self, "opacity_value_var"):
            self.opacity_value_var.set(
                "—"
                if not enabled
                else f"{100 - self.regions[self.selected_index].opacity}%"
            )
        if hasattr(self, "border_width_value_var"):
            self.border_width_value_var.set(
                "—"
                if not enabled
                else f"{self.regions[self.selected_index].border_width} px"
            )

    def choose_color(self) -> None:
        region = self._selected_region()
        if region is None:
            return
        color = colorchooser.askcolor(color=region.color, title=self._t("region_color"))[1]
        if color:
            region.color = color
            self._refresh_regions()
            self._render()

    def rename_region(self) -> None:
        region = self._selected_region()
        if region is None:
            return
        name = simpledialog.askstring(
            self._t("rename_region"),
            self._t("name"),
            initialvalue=region.name,
            parent=self.root,
        )
        if name and name.strip():
            region.name = name.strip()
            self._refresh_regions()
            self._update_statistics()

    def delete_selected_region(self, _event: tk.Event | None = None) -> None:
        if self.selected_index is None:
            return
        if self.drawing:
            self._finish_lasso()
        removed_index = self.selected_index
        del self.regions[removed_index]
        if self.base_region_index == removed_index:
            self.base_region_index = None
        elif self.base_region_index is not None and self.base_region_index > removed_index:
            self.base_region_index -= 1
        self.selected_index = (
            min(removed_index, len(self.regions) - 1) if self.regions else None
        )
        self.selected_vertex = None
        self.dragging_vertex = None
        self.hovered_vertex = None
        if self.selected_index is None and self.delete_vertex_mode:
            self._set_delete_vertex_mode(False)
        self._sync_region_controls()
        self._update_statistics()
        self._refresh_regions()
        self._render()
        self.status_var.set(self._t("region_deleted"))

    def _delete_selected_region(self, event: tk.Event) -> None:
        if event.widget is self.region_list:
            self.delete_selected_region()

    def delete_selected_vertex(self) -> None:
        target = self.hovered_vertex
        if target is None and self.selected_index is not None and self.selected_vertex is not None:
            target = (self.selected_index, self.selected_vertex)
        if target is None:
            messagebox.showinfo(
                self._t("vertex_not_selected"), self._t("select_vertex")
            )
            return
        region_index, vertex_index = target
        if not (0 <= region_index < len(self.regions)):
            self.hovered_vertex = None
            self.selected_vertex = None
            return
        region = self.regions[region_index]
        if len(region.points) <= 3:
            messagebox.showinfo(
                self._t("minimum_points"), self._t("keep_three_points")
            )
            return
        del region.points[vertex_index]
        if self.drawing and region_index == self.lasso_region_index:
            self.lasso_points = region.points
        self.selected_index = region_index
        self.selected_vertex = None
        self.hovered_vertex = None
        self.dragging_vertex = None
        self._update_statistics()
        self._refresh_regions()
        self._set_canvas_cursor()
        self._render()

    def _selected_region(self) -> Region | None:
        if self.selected_index is None:
            messagebox.showinfo(
                self._t("region_not_selected"), self._t("select_region")
            )
            return None
        return self.regions[self.selected_index]

    def _choose_export_regions(
        self,
    ) -> tuple[list[int], bool, bool] | None:
        if not self.regions:
            return [], True, False

        dialog = tk.Toplevel(self.root)
        dialog.title(self._t("export_options_title"))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=self._t("export_regions_prompt")).pack(anchor=tk.W)

        selected_vars = [tk.BooleanVar(value=True) for _ in self.regions]
        regions_frame = ttk.Frame(frame)
        regions_frame.pack(fill=tk.X, pady=(8, 12))
        for index, (region, variable) in enumerate(zip(self.regions, selected_vars)):
            ttk.Checkbutton(
                regions_frame, text=region.name, variable=variable
            ).grid(row=index, column=0, sticky=tk.W, pady=2)

        selection_buttons = ttk.Frame(frame)
        selection_buttons.pack(fill=tk.X, pady=(0, 10))

        def set_all(value: bool) -> None:
            for variable in selected_vars:
                variable.set(value)

        ttk.Button(
            selection_buttons,
            text=self._t("select_all"),
            command=lambda: set_all(True),
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            selection_buttons, text=self._t("select_none"), command=lambda: set_all(False)
        ).pack(side=tk.LEFT)

        include_intersections = tk.BooleanVar(value=True)
        crop_to_regions = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame,
            text=self._t("include_intersections"),
            variable=include_intersections,
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            frame,
            text=self._t("crop_export"),
            variable=crop_to_regions,
        ).pack(anchor=tk.W, pady=(4, 12))

        result: dict[str, tuple[list[int], bool, bool] | None] = {"value": None}

        def accept() -> None:
            indices = [
                index for index, variable in enumerate(selected_vars) if variable.get()
            ]
            if not indices:
                messagebox.showinfo(
                    self._t("export_options_title"),
                    self._t("select_export_region"),
                    parent=dialog,
                )
                return
            result["value"] = (
                indices,
                include_intersections.get(),
                crop_to_regions.get(),
            )
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill=tk.X)
        ttk.Button(
            buttons, text=self._t("cancel"), command=dialog.destroy
        ).pack(side=tk.RIGHT)
        ttk.Button(
            buttons, text=self._t("continue"), command=accept
        ).pack(side=tk.RIGHT, padx=(0, 6))
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.wait_window()
        return result["value"]

    def export_image(self) -> None:
        if self.image is None:
            messagebox.showinfo(self._t("no_image"), self._t("load_first"))
            return
        options = self._choose_export_regions()
        if options is None:
            return
        selected_indices, include_intersections, crop_to_regions = options
        path = filedialog.asksaveasfilename(
            title=self._t("image_export_title"),
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg *.jpeg"), ("TIFF", "*.tif *.tiff")],
        )
        if not path:
            return
        dpi = int(self.dpi_var.get())
        result = make_region_export(
            self.image,
            self.regions,
            selected_indices,
            include_intersections=include_intersections,
            crop_to_regions=crop_to_regions,
        )
        try:
            result.save(path, dpi=(dpi, dpi))
        except (OSError, ValueError) as exc:
            messagebox.showerror(self._t("export_error"), str(exc))
            return
        self.status_var.set(
            self._t(
                "image_saved",
                name=Path(path).name,
                width=result.width,
                height=result.height,
                dpi=dpi,
            )
        )

    def export_excel(self) -> None:
        if self.image is None:
            messagebox.showinfo(self._t("no_image"), self._t("load_first"))
            return
        path = filedialog.asksaveasfilename(
            title=self._t("table_export_title"),
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return

        base_pixels = self._base_area_pixels()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = self._t("excel_sheet")
        sheet.append(
            [self._t("excel_image"), self.image_path.name if self.image_path else ""]
        )
        sheet.append(
            [self._t("excel_size"), f"{self.image.width} × {self.image.height}"]
        )
        sheet.append([self._t("excel_base_area"), base_pixels])
        sheet.append([self._t("excel_base"), self._base_area_label()])
        sheet.append([])
        sheet.append(
            [
                self._t("excel_region"),
                self._t("excel_region_area"),
                self._t("excel_base_share"),
                self._t("excel_color"),
            ]
        )
        for cell in sheet[6]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
        for region, area in zip(self.regions, self.region_areas):
            percentage = calculate_area_percentage(area, base_pixels)
            sheet.append(
                [
                    region.name,
                    area,
                    None if percentage is None else round(percentage, 4),
                    region.color,
                ]
            )
            sheet.cell(sheet.max_row, 4).fill = PatternFill(
                "solid", fgColor=region.color.lstrip("#")
            )

        sheet.append([])
        sheet.append([self._t("excel_sum_overlap_area"), sum(self.region_areas)])
        total_percentage = calculate_area_percentage(sum(self.region_areas), base_pixels)
        sheet.append(
            [
                self._t("excel_sum_overlap_percent"),
                None if total_percentage is None else round(total_percentage, 4),
            ]
        )
        sheet.append([self._t("excel_union_area"), self.union_area])
        union_percentage = calculate_area_percentage(self.union_area, base_pixels)
        sheet.append(
            [
                self._t("excel_union_percent"),
                None if union_percentage is None else round(union_percentage, 4),
            ]
        )
        sheet.column_dimensions["A"].width = 54
        sheet.column_dimensions["B"].width = 24
        sheet.column_dimensions["C"].width = 24
        sheet.column_dimensions["D"].width = 14
        try:
            workbook.save(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror(self._t("export_error"), str(exc))
            return
        self.status_var.set(self._t("table_saved", name=Path(path).name))


def main() -> None:
    root = tk.Tk()
    ImageAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
