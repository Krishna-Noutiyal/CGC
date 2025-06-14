import flet as ft
from config import ColorScheme
from scripts import CSVProcessor, ExcelProcessor  # type: ignore
import os


class MainView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.csv_processor = CSVProcessor()
        self.selected_files = []
        self.output_path = ""

        # UI Components
        self.file_picker = ft.FilePicker(on_result=self.on_files_selected)
        self.output_picker = ft.FilePicker(on_result=self.on_output_selected)
        self.page.overlay.extend([self.file_picker, self.output_picker])

        self.selected_files_text = ft.Text(
            "No CSVs Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.output_path_text = ft.Text(
            "No Output Path Selected", color=ColorScheme.TEXT_SECONDARY, size=14
        )

        self.status_text = ft.Text("", color=ColorScheme.TEXT_SECONDARY, size=14)

    def on_files_selected(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.selected_files = [file.path for file in e.files]
            file_names = [os.path.basename(path) for path in self.selected_files]
            self.selected_files_text.value = (
                f"Selected {len(self.selected_files)} Files: {', '.join(file_names)}"
            )
            self.selected_files_text.color = ColorScheme.SUCCESS
        else:
            self.selected_files = []
            self.selected_files_text.value = "No CSVs Selected"
            self.selected_files_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    def on_output_selected(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.output_path = e.path
            self.output_path_text.value = (
                f"Output: {os.path.basename(self.output_path)}"
            )
            self.output_path_text.color = ColorScheme.SUCCESS
        else:
            self.output_path = ""
            self.output_path_text.value = "No Output Path Selected"
            self.output_path_text.color = ColorScheme.TEXT_SECONDARY
        self.page.update()

    def on_submit_clicked(self, e):
        if not self.selected_files:
            self.show_status("Please Select CSVs !", ColorScheme.ERROR)
            return

        if not self.output_path:
            self.show_status("Please Select Output Path !", ColorScheme.ERROR)
            return

        try:
            self.show_status("Processing Files...", ColorScheme.PRIMARY)

            # Combine CSV files into a single DataFrame
            dataframe = self.csv_processor.combine_csvs(self.selected_files)

            create_Excel = ExcelProcessor(df=dataframe).Make_Excel(self.output_path)

            if create_Excel:
                self.show_status(
                    "Excel File Created Successfully !", ColorScheme.SUCCESS
                )
            else:
                self.show_status("Error Processing Files !", ColorScheme.ERROR)
        except Exception as ex:
            self.show_status(f"Error: {str(ex)}", ColorScheme.ERROR)

    def show_status(self, message: str, color: str):
        self.status_text.value = message
        self.status_text.color = color
        self.status_text.weight = ft.FontWeight.BOLD

        self.page.update()

    def build(self):
        return ft.Container(
            content=ft.Column(
                [
                    # Title
                    ft.Container(
                        content=ft.Text(
                            "CGC : Capital Gain Calculator",
                            size=32,
                            weight=ft.FontWeight.BOLD,
                            color=ColorScheme.PRIMARY,
                        ),
                        margin=ft.margin.only(bottom=10),
                    ),
                    # Description
                    ft.Container(
                        content=ft.Text(
                            "Hello, CGC creates beautiful excel dashboard for visualizing Capital Gain Please select the Capital Gain CSVs download from the AIS portal \nGive the Output Path to store the Capital Gain (.xlsx) File",
                            size=16,
                            color=ColorScheme.TEXT_SECONDARY,
                        ),
                        margin=ft.margin.only(bottom=30),
                    ),
                    # File Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Select CSV Files:",
                                    size=18,
                                    weight=ft.FontWeight.W_500,
                                    color=ColorScheme.TEXT_PRIMARY,
                                ),
                                ft.Container(
                                    content=ft.Row(
                                        [
                                            ft.ElevatedButton(
                                                "Browse Files",
                                                icon=ft.Icons.FOLDER_OPEN,
                                                on_click=lambda _: self.file_picker.pick_files(
                                                    allow_multiple=True,
                                                    allowed_extensions=["csv"],
                                                ),
                                                bgcolor=ColorScheme.PRIMARY,
                                                color=ft.Colors.WHITE,
                                                width=200,
                                                height=50,
                                                style=ft.ButtonStyle(
                                                    text_style=ft.TextStyle(
                                                        size=16,
                                                        weight=ft.FontWeight.BOLD,
                                                    )  # Increased text size
                                                ),
                                            )
                                        ]
                                    ),
                                    margin=ft.margin.only(top=5, bottom=10),
                                ),
                                self.selected_files_text,
                            ]
                        ),
                        padding=20,
                        border=ft.border.all(1, ColorScheme.BORDER),
                        border_radius=8,
                        bgcolor=ColorScheme.SURFACE,
                        margin=ft.margin.only(bottom=20),
                    ),
                    # Output Path Selection Section
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Output Path:",
                                    size=18,
                                    weight=ft.FontWeight.W_500,
                                    color=ColorScheme.TEXT_PRIMARY,
                                ),
                                ft.Container(
                                    content=ft.Row(
                                        [
                                            ft.ElevatedButton(
                                                "Output Path",
                                                icon=ft.Icons.SAVE,
                                                on_click=lambda _: self.output_picker.save_file(
                                                    file_name="Capital Gain.xlsx",
                                                    allowed_extensions=["xlsx"],
                                                ),
                                                bgcolor=ColorScheme.SECONDARY,
                                                color=ColorScheme.TEXT_PRIMARY,
                                                width=200,
                                                height=50,
                                                style=ft.ButtonStyle(
                                                    text_style=ft.TextStyle(
                                                        size=16,
                                                        weight=ft.FontWeight.BOLD,
                                                    )  # Increased text size
                                                ),
                                            )
                                        ]
                                    ),
                                    margin=ft.margin.only(top=5, bottom=10),
                                ),
                                self.output_path_text,
                            ]
                        ),
                        padding=20,
                        border=ft.border.all(1, ColorScheme.BORDER),
                        border_radius=8,
                        bgcolor=ColorScheme.SURFACE,
                        margin=ft.margin.only(bottom=30),
                    ),
                    # Submit Button
                    ft.Container(
                        content=ft.ElevatedButton(
                            "Submit",
                            icon=ft.Icons.PLAY_ARROW,
                            on_click=self.on_submit_clicked,
                            bgcolor=ColorScheme.SUCCESS,
                            color=ft.Colors.WHITE,
                            width=200,
                            height=50,
                        ),
                        alignment=ft.alignment.center,
                        margin=ft.margin.only(bottom=20),
                    ),
                    # Status Text
                    ft.Container(
                        content=self.status_text, alignment=ft.alignment.center
                    ),
                ]
            ),
            bgcolor=ColorScheme.BACKGROUND,
            padding=50,
            expand=True,
            border_radius=15
        )
