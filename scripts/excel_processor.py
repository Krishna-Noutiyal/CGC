from scripts.csv_processor import CSVProcessor
import glob, os
import pandas as pd
from dataclasses import dataclass
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


@dataclass
class ExcelProcessor:

    df: pd.DataFrame

    def _create_workbook(self, file_name: str) -> None:
        """
        Create an Excel workbook with the given file name.
        """
        self.workbook = xlsxwriter.Workbook(file_name)

    def _create_worksheet(self, sheet_name: str) -> None:
        """
        Create a worksheet in the workbook with the given sheet name.
        """
        self.worksheet = self.workbook.add_worksheet(sheet_name)

    def _add_formats(self) -> dict:
        """
        Add multiple formats to the workbook and return them as a dictionary.
        The formatting is for dark mode excel sheets.
        """
        base_format = {"align": "center", "valign": "vcenter", "text_wrap": True}
        formats = {
            "super": self.workbook.add_format(
                {
                    **base_format,
                    "font_script": 1,
                }
            ),
            "blank": self.workbook.add_format(
                {
                    **base_format,
                    "font_size": 14,
                }
            ),
            "orange_h": self.workbook.add_format(
                {**base_format, "bold": True, "bg_color": "#E97132", "font_size": 18}
            ),
            "blue_h": self.workbook.add_format(
                {**base_format, "bg_color": "#4D93D9", "font_size": 16}
            ),
            "dblue_h": self.workbook.add_format(
                {**base_format, "bg_color": "#83CCEB", "font_size": 16}
            ),
            "red_h": self.workbook.add_format(
                {**base_format, "bg_color": "#FF7979", "font_size": 16}
            ),
            "green_h": self.workbook.add_format(
                {**base_format, "bg_color": "#00B050", "font_size": 16}
            ),
            "grey_h": self.workbook.add_format(
                {**base_format, "bold": True, "bg_color": "#DAE9F8", "font_size": 18}
            ),
            "black_h": self.workbook.add_format(
                {**base_format, "bold": True, "font_size": 26}
            ),
        }
        return formats

    def _set_cell_dimensions(self, width: float = 34.5, height: float = 48.3) -> None:
        """
        Set the width and height for all columns and rows in the worksheet.
        """
        # Set default column width and row height for the worksheet
        self.worksheet.set_default_row(height)
        self.worksheet.set_column(0, self.df.shape[1] - 1, width)

    def Make_Excel(self, file_path: str) -> bool:
        """
        Create an Excel file with the given file name and write the DataFrame to it.
        """
        print(f"Creating Excel file at {file_path}...")

        try:
            self._create_workbook(file_path)

            """ ##################### Capital Gains Dashboard ##################### """

            self._create_worksheet("Capital Gains")
            # Set cell dimensions for the data worksheet
            self._set_cell_dimensions()

            formats = self._add_formats()

            # Set cell dimensions
            self._set_cell_dimensions()

            """ ##################### FORMATING CELLS ##################### """
            # Creating Dashboard similar to Dashboardv2.xlsx in the Dashboards folder

            print("Formatting cells...")

            ws = self.worksheet

            # Border format for the cells
            # This is used to create a border around the cells
            # border_format = self.workbook.add_format({'border': 1})
            # for row in range(12):
            #     for col in range(6):
            #         ws.write_blank(row, col, None, border_format)

            # Row and column starts from 0 insted of 1

            # This is the 1st row
            ws.merge_range(0, 1, 0, 5, "SHORT TERM", formats["orange_h"])
            # This is the 2nd row
            ws.write(1, 1, "Full Value of Consideration", formats["blue_h"])
            ws.write(1, 2, "Cost of Acquisition", formats["dblue_h"])
            ws.write(1, 3, "Tax", formats["red_h"])
            ws.merge_range(1, 4, 1, 5, "Short Term Tax", formats["grey_h"])

            # Short Term Before 23rd July 2024
            # Create a superscript "rd" only, rest normal
            ws.write_rich_string(
                2,
                0,
                formats["blue_h"],
                "Before 23",
                formats["super"],
                "rd",
                formats["blue_h"],
                " July 2024",
                formats["blue_h"],
            )

            # Short Term After 23rd July 2024
            # Create a superscript "rd" only, rest normal
            ws.write_rich_string(
                3,
                0,
                formats["dblue_h"],
                "After 23",
                formats["super"],
                "rd",
                formats["dblue_h"],
                " July 2024",
                formats["dblue_h"],
            )

            # Short Term Grand Total
            ws.write(4, 0, "Grand Total", formats["green_h"])

            # Long Term Section
            ws.merge_range(5, 1, 5, 5, "LONG TERM", formats["orange_h"])
            ws.write(6, 1, "Full Value of Consideration", formats["blue_h"])
            ws.write(6, 2, "Cost of Acquisition", formats["dblue_h"])
            ws.write(6, 3, "Tax", formats["red_h"])
            ws.merge_range(6, 4, 6, 5, "Long Term Tax", formats["grey_h"])

            # Long Term Before 23rd July 2024
            # Create a superscript "rd" only, rest normal
            ws.write_rich_string(
                7,
                0,
                formats["blue_h"],
                "Before 23",
                formats["super"],
                "rd",
                formats["blue_h"],
                " July 2024",
                formats["blue_h"],
            )

            # Long Term After 23rd July 2024
            # Create a superscript "rd" only, rest normal
            ws.write_rich_string(
                8,
                0,
                formats["dblue_h"],
                "After 23",
                formats["super"],
                "rd",
                formats["dblue_h"],
                " July 2024",
                formats["dblue_h"],
            )

            # Long Term Grand Total
            ws.write(9, 0, "Grand Total", formats["green_h"])

            print("Calculating values...")

            """ ##################### CALCULATING VALUES ##################### """
            before_23 = self.df[
                self.df["Date of Sale/Transfer"] < pd.to_datetime("2024-07-23")
            ]
            after_23 = self.df[
                self.df["Date of Sale/Transfer"] >= pd.to_datetime("2024-07-23")
            ]

            # Filter for "Short term" asset type before and after 23rd July 2024
            short_before_23 = before_23[before_23["Asset Type"] == "Short term"]
            short_after_23 = after_23[after_23["Asset Type"] == "Short term"]

            # Filter for "Long term" asset type before and after 23rd July 2024
            long_before_23 = before_23.drop(short_before_23.index)
            long_after_23 = after_23.drop(short_after_23.index)

            fvc_sort_Before_23 = short_before_23[
                "Sales Consideration - Reported by Source"
            ].sum()
            fvc_sort_After_23 = short_after_23[
                "Sales Consideration - Reported by Source"
            ].sum()
            fvc_long_Before_23 = long_before_23[
                "Sales Consideration - Reported by Source"
            ].sum()
            fvc_long_After_23 = long_after_23[
                "Sales Consideration - Reported by Source"
            ].sum()
            fvc_long_After_23 = long_after_23[
                "Sales Consideration - Reported by Source"
            ].sum()

            coa_short_Before_23 = short_before_23["Cost of Acquisition"].sum()
            coa_short_After_23 = short_after_23["Cost of Acquisition"].sum()
            coa_long_Before_23 = long_before_23["Cost of Acquisition"].sum()
            coa_long_After_23 = long_after_23["Cost of Acquisition"].sum()

            # Profit/Loss calculations in short and long term
            # Short term profit/loss is calculated as Full Value of Consideration - Cost of Acquisition
            pnl_short = (fvc_sort_Before_23 + fvc_sort_After_23) - (
                coa_short_Before_23 + coa_short_After_23
            )
            pnl_long = (fvc_long_Before_23 + fvc_long_After_23) - (
                coa_long_Before_23 + coa_long_After_23
            )

            print("Inserting data into the worksheet...")

            """ ##################### INSERTING DATA INTO THE WORKSHEET ##################### """

            """ ##################### SHORT TERM VALUES ##################### """
            # Short Term Before 23rd July 2024 Values
            ws.write_number(2, 1, fvc_sort_Before_23, formats["blank"])
            ws.write_number(2, 2, coa_short_Before_23, formats["blank"])
            ws.write_formula(
                2, 3, "=IF(B3-C3>=0,ROUND((B3-C3)*15%,0),0)", formats["blank"]
            )

            # Short Term After 23rd July 2024 Values
            ws.write_number(3, 1, fvc_sort_After_23, formats["blank"])
            ws.write_number(3, 2, coa_short_After_23, formats["blank"])
            ws.write_formula(
                3, 3, "=IF(B4-C4>=0,ROUND((B4-C4)*20%,0),0)", formats["blank"]
            )

            # Short Term Grand Total Values
            ws.write_formula(4, 1, "=SUM(B3:B4)", formats["green_h"])
            ws.write_formula(4, 2, "=SUM(C3:C4)", formats["green_h"])
            ws.write_formula(4, 3, "=SUM(D3:D4)", formats["green_h"])

            # Short Term Tax
            ws.merge_range(2, 4, 4, 5, "=D5", formats["black_h"])

            """ ##################### LONG TERM VALUES ##################### """

            # long Term Before 23rd July 2024 Values
            ws.write_number(7, 1, fvc_long_Before_23, formats["blank"])
            ws.write_number(7, 2, coa_long_Before_23, formats["blank"])
            ws.write_formula(
                7,
                3,
                "=IF(B8-C8>100000,ROUND(((B8-C8)-100000)*10%,0),0)",
                formats["blank"],
            )

            # long Term After 23rd July 2024 Values
            ws.write_number(8, 1, fvc_long_After_23, formats["blank"])
            ws.write_number(8, 2, coa_long_After_23, formats["blank"])
            ws.write_formula(
                8,
                3,
                "=IF(B9-C9>125000,ROUND(((B9-C9)-125000)*12.5%,0),0)",
                formats["blank"],
            )

            # long Term Grand Total Values
            ws.write_formula(9, 1, "=SUM(B8:B9)", formats["green_h"])
            ws.write_formula(9, 2, "=SUM(C8:C9)", formats["green_h"])
            ws.write_formula(9, 3, "=SUM(D8:D9)", formats["green_h"])

            # long Term Tax
            ws.merge_range(7, 4, 9, 5, "=D10", formats["black_h"])

            """ ##################### GRAND TOTAL OF TAX ##################### """

            # Grand Total Tax at bottom
            ws.merge_range(10, 0, 10, 1, "Total Tax", formats["grey_h"])
            ws.merge_range(10, 2, 10, 4, "=SUM(E3,E8)", formats["black_h"])

            """ ##################### Short n Long Term Profit/Loss ##################### """

            # Write the Short Term and Long Term Profit/Loss at the bottom
            if pnl_short >= 0:
                ws.write(11, 0, "Short Term Profit", formats["green_h"])
            elif pnl_short < 0:
                ws.write(11, 0, "Short Term Loss", formats["red_h"])

            if pnl_long >= 0:
                ws.write(11, 3, "Long Term Profit", formats["green_h"])
            elif pnl_long < 0:
                ws.write(11, 3, "Long Term Loss", formats["red_h"])

            # Write the values of Short Term and Long Term Profit/Loss
            ws.merge_range(11, 1, 11, 2, "=B5-C5", formats["black_h"])
            ws.merge_range(11, 4, 11, 5, "=B10-C10", formats["black_h"])

            """ ##################### Capital Gains Data ##################### """

            # Create a new worksheet for the raw DataFrame
            self._create_worksheet("Capital Gains Data")

            # Set cell dimensions for the data worksheet
            self._set_cell_dimensions()
            data_ws = self.worksheet
            data_formats = self._add_formats()

            # Set header row height and column widths
            data_ws.set_row(0, 55)
            data_ws.set_column(0, self.df.shape[1] - 1, 26)
            # Set data row heights (from row 1 onwards)
            for row_num in range(1, len(self.df) + 2):  # +2 to include totals row
                data_ws.set_row(row_num, 30)

                # Write the header with formatting (font size 16)
                header_format = self.workbook.add_format(
                    {
                        **{
                            "align": "center",
                            "valign": "vcenter",
                            "text_wrap": True,
                            "bold": True,
                            "bg_color": "#E97132",
                            "font_size": 16,
                        }
                    }
                )
                for col_num, col_name in enumerate(self.df.columns):
                    data_ws.write(0, col_num, col_name, header_format)

            # Write the data rows
            for row_num, row in enumerate(self.df.itertuples(index=False), start=1):
                for col_num, value in enumerate(row):
                    # Set font size 11 for the first cell (first column, first data row)
                    if col_num == 0:
                        fmt = self.workbook.add_format(
                            {
                                "align": "center",
                                "valign": "vcenter",
                                "text_wrap": True,
                                "font_size": 11,
                            }
                        )
                    else:
                        fmt = data_formats["blank"]
                    if isinstance(value, (int, float)):
                        data_ws.write_number(row_num, col_num, value, fmt)
                    else:
                        data_ws.write(row_num, col_num, value, fmt)

            # Write totals for numeric columns at the end
            total_row = len(self.df) + 1
            for col_num, col_name in enumerate(self.df.columns):
                col_letter = xl_col_to_name(col_num)
                formula = f"=SUM({col_letter}2:{col_letter}{len(self.df)+1})"
                data_ws.write_formula(
                    total_row, col_num, formula, data_formats["green_h"]
                )
            # Write "Total" label in the first column of the totals row
            data_ws.write(total_row, 0, "Total", data_formats["grey_h"])

            # self.worksheet.activate()
            self.workbook.close()
            print(f"Excel file created successfully at {file_path}")

            return True
        except Exception as e:
            print(f"Error creating Excel file: {str(e)}")
            return False


if __name__ == "__main__":
    # Create an instance of CSVProcessor
    test = CSVProcessor()

    # List all CSV files in the /test folder
    test_folder = "short_sale_calculator/test"
    file_list = glob.glob(os.path.join(test_folder, "*.csv"))
    print("Files found:", file_list)

    # Example usage of combine_csvs with the found files
    df = test.combine_csvs(file_list)

    summary = ExcelProcessor(df=df)

    summary.Make_Excel("short_sale_calculator/test/Capital_Gains_Summary.xlsx")
