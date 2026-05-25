import re
import pandas as pd
from pathlib import Path
from typing import List, Optional, Tuple
import matplotlib.pyplot as plt


def infer_country_code(file_name: str) -> str:
    """Extract a country code from the filename."""
    file_name = Path(file_name).name
    if match := re.match(r"^([A-Za-z]{2})_[0-9]{8}_", file_name):
        return match.group(1).upper()
    basename = Path(file_name).stem
    if matches := re.findall(r"(?:^|_)([A-Za-z]{2})(?=_|$)", basename):
        return matches[0].upper()
    return basename[:2].upper()


def sort_files_by_country(file_names: List[str]) -> List[str]:
    grouped = {}
    for name in file_names:
        country = infer_country_code(name)
        grouped.setdefault(country, []).append(name)
    ordered = []
    for country in sorted(grouped):
        ordered.extend(sorted(grouped[country]))
    return ordered


def build_y_labels(file_names: List[str]) -> List[str]:
    return [f"{infer_country_code(name):<3} {name}" for name in file_names]

# Data class for reading and comparing cell values from an Excel sheet
class CountrySheet:
    def __init__(self, file_path: str | Path, country_code: str = "pt") -> None:
        """Initialize the sheet with file path and prepare error list."""
        self.FilePath: Path = Path(file_path)
        self.errors: List[str] = []
        self._sheet_name: Optional[str | int] = None
        self._cached_df: Optional[pd.DataFrame] = None

    def _looks_like_new_format(self) -> bool:
        return bool(re.match(r"^[A-Za-z]{2}_[0-9]{8}_", self.FilePath.name))

    def _load_data_frame(self) -> Optional[pd.DataFrame]:
        if self._cached_df is not None:
            return self._cached_df

        try:
            xls = pd.ExcelFile(self.FilePath)
        except Exception as exc:
            self.errors.append(f"File could not be read: {exc}")
            return None

        candidates = []
        if self._looks_like_new_format() and "export" in xls.sheet_names:
            candidates.append("export")
        if "export" in xls.sheet_names and "export" not in candidates:
            candidates.append("export")
        if "data" in xls.sheet_names:
            candidates.append("data")
        candidates.extend(xls.sheet_names)

        for sheet in candidates:
            try:
                df = pd.read_excel(self.FilePath, sheet_name=sheet)
            except Exception:
                continue
            first_col = df.iloc[:, 0].astype(str).str.strip().str.lower()
            if (first_col == "pt").any():
                self._sheet_name = sheet
                self._cached_df = df
                return df

        try:
            self._sheet_name = xls.sheet_names[0]
            self._cached_df = pd.read_excel(self.FilePath, sheet_name=0)
            return self._cached_df
        except Exception as exc:
            self.errors.append(f"File could not be read: {exc}")
            return None

    def migrant_native_ratio(
        self,
        row: str,
        col1: str = "na",
        col2: str = "ma",
        normalize: bool = False
    ) -> Optional[Tuple[float, float, float]]:
        """Compare migrant and native values in a specific row and return the ratio.

        Returns a tuple(ratio, native_value, migrant_value) or None if an error occurred.
        """
        df = self._load_data_frame()
        if df is None:
            return None

        cleaned_first_col = df.iloc[:, 0].astype(str).str.strip()
        search_row = df[cleaned_first_col == str(row).strip()]
        if search_row.empty:
            self.errors.append(f"Row with '{row}' not found")
            return None

        def find_column(prefix: str):
            prefix_lower = prefix.lower()
            for col in df.columns:
                if str(col).strip().lower().startswith(prefix_lower):
                    return col
            if not df.empty:
                header_codes = [str(x).strip().lower() for x in df.iloc[0].tolist()]
                for idx, code in enumerate(header_codes):
                    if code and code != "nan" and code.startswith(prefix_lower):
                        return df.columns[idx]
            return None

        native_col = find_column(col1)
        migrant_col = find_column(col2)

        if native_col is None:
            self.errors.append(f"No column with prefix '{col1}' found")
        if migrant_col is None:
            self.errors.append(f"No column with prefix '{col2}' found")
        if native_col is None or migrant_col is None:
            return None

        try:
            native_value = float(search_row[native_col].iloc[0])
            if pd.isna(native_value):
                raise ValueError("native value is missing")
        except Exception as exc:
            self.errors.append(f"Native value could not be read or converted: {exc}")
            return None

        try:
            migrant_value = float(search_row[migrant_col].iloc[0])
            if pd.isna(migrant_value):
                raise ValueError("migrant value is missing")
        except Exception as exc:
            self.errors.append(f"Migrant value could not be read or converted: {exc}")
            return None

        if native_value == 0:
            self.errors.append("Native value is 0, ratio cannot be calculated")
            return None

        ratio = migrant_value / native_value
        return ratio, native_value, migrant_value

    def print_summary(self) -> None:
        """Print a formatted summary of errors."""
        if self.errors:
            print("Errors:")
            for error in self.errors:
                print(f"  - {error}")
        else:
            print("Errors: none")

# Main execution
if __name__ == "__main__":
    data_dir = Path(__file__).resolve().parent / "data"
    cache_file = data_dir / "plausibility_results.csv"

    if cache_file.exists():
        print(f"Loading cached results from {cache_file}")
        result_df = pd.read_csv(cache_file)
    else:
        print("Computing results...")
        excel_files = sorted(
            [f for f in data_dir.glob("*.xlsx") if not f.name.startswith("~$")]
        )

        # Migrant analysis
        # ------------------------------------------------------------
        results = []
        if excel_files:
            total_files = len(excel_files)
            for idx, file_path in enumerate(excel_files, start=1):
                print(f"Processing file {idx}/{total_files}: {file_path.name}")
                sheet = CountrySheet(file_path)
                try:
                    # Extract population number for normalization
                    result_tuple = sheet.migrant_native_ratio(row="pt")
                    if result_tuple is None:
                        continue
                    try:
                        ratio, native_count, migrant_count = result_tuple
                    except Exception:
                        continue
                    # Extract females number for normalization
                    result_tuple = sheet.migrant_native_ratio(row="pt", col1="nf", col2="mf")
                    if result_tuple is None:
                        continue
                    try:
                        ratio, native_females, migrant_females = result_tuple
                    except Exception:
                        continue
                    # Extract other values
                    for row_label, label in [
                    ("pt", "populations"),
                    ("bt", "births"),
                    ("f", "fertility (official)"),
                    ("fc", "fertility (calculated)"),
                    ("dt", "deaths"),
                    ("mt", "migrations"),
                    ("et", "emigrations"),
                    ("l", "life expectancy")
                ]:
                        result_tuple = sheet.migrant_native_ratio(row=row_label)
                        if result_tuple is None:
                            continue
                        try:
                            ratio, native_value, migrant_value = result_tuple
                        except Exception:
                            continue
                        # Normalization
                        if row_label == "pt":
                            native_value = 100.0*native_value/(native_count + migrant_count)
                            migrant_value = 100.0*migrant_value/(native_count + migrant_count)
                        elif row_label == "bt":
                            try:
                                native_value = 1000000*native_value/native_females
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                native_value = None
                            try:
                                migrant_value = 1000000*migrant_value/migrant_females
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                migrant_value = None
                        elif row_label in ["dt", "mt", "et"]:
                            try:
                                native_value = 100.0*native_value/native_count
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                native_value = None
                            try:
                                migrant_value = 100.0*migrant_value/migrant_count
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                migrant_value = None
                        results.append({
                            "file_name": file_path.name,
                            "row_label": row_label,
                            "label": label,
                            "native_value": native_value,
                            "migrant_value": migrant_value
                        })
                finally:
                    sheet.print_summary()

        result_df = pd.DataFrame(
            results,
            columns=["file_name", "row_label", "label", "native_value", "migrant_value"],
        )
        # Cache the DataFrame to file
        result_df.to_csv(cache_file, index=False)
        print(f"Results saved to {cache_file}")

        # Native analysis
        # ------------------------------------------------------------
        native_results = []
        if excel_files:
            total_files = len(excel_files)
            for idx, file_path in enumerate(excel_files, start=1):
                print(f"Processing file {idx}/{total_files} (Native): {file_path.name}")
                sheet = CountrySheet(file_path)
                try:
                    # Extract population number for normalization (native only)
                    result_tuple = sheet.migrant_native_ratio(row="pt")
                    if result_tuple is None:
                        continue
                    try:
                        _, native_count, _ = result_tuple
                    except Exception:
                        continue
                    # Extract females number for normalization (native only)
                    result_tuple = sheet.migrant_native_ratio(row="pt", col1="nf", col2="mf")
                    if result_tuple is None:
                        continue
                    try:
                        _, native_females, _ = result_tuple
                    except Exception:
                        continue
                    for row_label, label in [
                    ("pt", "populations"),
                    ("bt", "births"),
                    ("f", "fertility (official)"),
                    ("fc", "fertility (calculated)"),
                    ("dt", "deaths"),
                    ("mt", "migrations"),
                    ("et", "emigrations"),
                    ("l", "life expectancy")
                ]:
                        result_tuple = sheet.migrant_native_ratio(row=row_label)
                        if result_tuple is None:
                            continue
                        try:
                            _, native_value, migrant_value = result_tuple
                        except Exception:
                            continue
                        # Normalization (native only)
                        if row_label == "pt":
                            total_pop = native_value + migrant_value if native_value is not None and migrant_value is not None else None
                            native_value = 100.0 * native_value / total_pop if total_pop else None
                        elif row_label == "bt":
                            try:
                                native_value = 1000000*native_value/native_females
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                native_value = None
                        elif row_label in ["dt", "mt", "et"]:
                            try:
                                native_value = 100.0*native_value/native_count
                            except ZeroDivisionError:
                                print("Error: Cannot divide by zero!")
                                native_value = None
                        native_results.append({
                            "file_name": file_path.name,
                            "row_label": row_label,
                            "label": label,
                            "native_value": native_value
                        })
                finally:
                    sheet.print_summary()

        native_df = pd.DataFrame(
            native_results,
            columns=["file_name", "row_label", "label", "native_value"],
        )
        native_cache_file = data_dir / "plausibility_results_native.csv"
        native_df.to_csv(native_cache_file, index=False)
        print(f"Native results saved to {native_cache_file}")

    # Migrant plots
    # ------------------------------------------------------------
    plots_dir = Path(__file__).resolve().parent / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)

    labels_to_plot = [
        "populations",
        "births",
        "fertility (official)",
        "fertility (calculated)",
        "deaths",
        "migrations",
        "emigrations",
        "life expectancy",
    ]
    all_file_names = sort_files_by_country(result_df["file_name"].drop_duplicates().tolist())
    y_labels = build_y_labels(all_file_names)

    for label in labels_to_plot:
        df_label = result_df[result_df["label"] == label].copy()
        ratio_series = (
            df_label.set_index("file_name")["migrant_value"]
            .reindex(all_file_names)
        )
        mean_ratio = ratio_series.mean()
        plt.figure(figsize=(14, max(6, len(all_file_names) * 0.2)))
        bar_colors = ["steelblue" if not pd.isna(val) else "lightgray" for val in ratio_series]
        plt.barh(y_labels, ratio_series.fillna(0), color=bar_colors)
        for idx, val in enumerate(ratio_series):
            if pd.isna(val):
                plt.text(0, idx, "missing", va="center", ha="left", color="gray", fontsize=9)
        plt.axvline(
            mean_ratio,
            color="red",
            linestyle="--",
            linewidth=1.8,
            label=f"mean = {mean_ratio:.4f}",
        )
        ax = plt.gca()
        ax.set_yticklabels(y_labels, fontfamily="monospace")
        ax.invert_yaxis()
        plt.xlabel("Migrant value")
        plt.ylabel("Country / file_name")
        plt.title(f"{label.capitalize()} (Migrant)")
        plt.legend()
        plt.tight_layout()
        safe_label = label.replace(" ", "_").replace("(", "").replace(")", "")
        out_path = plots_dir / f"{safe_label}_migrant.png"
        plt.savefig(out_path, dpi=150)
        plt.show()
        plt.close()

    # Native plots
    # ------------------------------------------------------------
    native_all_file_names = sort_files_by_country(native_df["file_name"].drop_duplicates().tolist())
    native_y_labels = build_y_labels(native_all_file_names)
    for label in labels_to_plot:
        df_label = native_df[native_df["label"] == label].copy()
        native_series = (
            df_label.set_index("file_name")["native_value"]
            .reindex(native_all_file_names)
        )
        mean_native = native_series.mean()
        plt.figure(figsize=(14, max(6, len(native_all_file_names) * 0.2)))
        bar_colors = ["darkorange" if not pd.isna(val) else "lightgray" for val in native_series]
        plt.barh(native_y_labels, native_series.fillna(0), color=bar_colors)
        for idx, val in enumerate(native_series):
            if pd.isna(val):
                plt.text(0, idx, "missing", va="center", ha="left", color="gray", fontsize=9)
        plt.axvline(
            mean_native,
            color="red",
            linestyle="--",
            linewidth=1.8,
            label=f"mean = {mean_native:.4f}",
        )
        ax = plt.gca()
        ax.set_yticklabels(native_y_labels, fontfamily="monospace")
        ax.invert_yaxis()
        plt.xlabel("Native value")
        plt.ylabel("Country / file_name")
        plt.title(f"{label.capitalize()} (Native)")
        plt.legend()
        plt.tight_layout()
        safe_label = label.replace(" ", "_").replace("(", "").replace(")", "")
        out_path = plots_dir / f"{safe_label}_native.png"
        plt.savefig(out_path, dpi=150)
        plt.show()
        plt.close()

