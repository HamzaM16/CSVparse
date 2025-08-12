import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import re
import unicodedata


def search_excel_with_gui():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Prompt for file
    csv_file_path = filedialog.askopenfilename(
        title="Select CSV File", filetypes=[("CSV Files", "*.csv")]
    )

    if not csv_file_path:
        messagebox.showinfo("No File Selected", "No CSV file was selected.")
        return

    # Prompt for asset IDs (comma separated)
    asset_input = simpledialog.askstring(
        "Asset IDs", "Enter Asset IDs (comma separated):"
    )
    if not asset_input:
        messagebox.showinfo("Input Required", "Asset IDs are required.")
        return

    # Prompt for country codes (comma separated)
    country_input = simpledialog.askstring(
        "Country Codes", "Enter Country Codes (comma separated):"
    )
    if not country_input:
        messagebox.showinfo("Input Required", "Country codes are required.")
        return

    # Normalize inputs
    asset_ids = [
        asset_id.strip().upper()
        for asset_id in asset_input.split(",")
        if asset_id.strip()
    ]
    country_codes = [
        code.strip().upper() for code in country_input.split(",") if code.strip()
    ]

    try:
        # First, let's examine the raw file to detect the separator
        print("Examining raw file content...")

        with open(csv_file_path, "r", encoding="utf-8", errors="ignore") as f:
            first_line = f.readline().strip()
            second_line = f.readline().strip()

        print(f"First line: '{first_line}'")
        print(f"Second line: '{second_line}'")

        # Count potential separators in the header line
        separators = [",", "\t", ";", "|"]
        separator_counts = {}
        for sep in separators:
            count = first_line.count(sep)
            separator_counts[sep] = count
            print(f"'{sep}' appears {count} times in header")

        # Find the separator with the most occurrences
        best_separator = max(separator_counts, key=separator_counts.get)
        print(
            f"Best separator appears to be: '{best_separator}' (count: {separator_counts[best_separator]})"
        )

        # Try reading with the detected separator
        df = None
        if separator_counts[best_separator] > 0:
            try:
                if best_separator == "\t":
                    df = pd.read_csv(csv_file_path, sep="\t")
                else:
                    df = pd.read_csv(csv_file_path, sep=best_separator)
                print(
                    f"Successfully read with separator '{best_separator}': {df.shape}"
                )
                print(f"Columns: {list(df.columns)}")
            except Exception as e:
                print(f"Failed to read with detected separator: {e}")

        # If auto-detection failed, try each separator manually
        if df is None or len(df.columns) <= 1:
            print("Auto-detection failed, trying each separator manually...")

            for sep_name, sep_char in [
                ("comma", ","),
                ("tab", "\t"),
                ("semicolon", ";"),
                ("pipe", "|"),
            ]:
                try:
                    df = pd.read_csv(csv_file_path, sep=sep_char)
                    print(f"Trying {sep_name} separator: got {len(df.columns)} columns")
                    if len(df.columns) > 1:
                        print(f"Success with {sep_name}! Columns: {list(df.columns)}")
                        break
                    else:
                        print(
                            f"{sep_name} separator only gave 1 column: {list(df.columns)}"
                        )
                except Exception as e:
                    print(f"Failed with {sep_name}: {e}")

        # If still no success, try Excel engine
        if df is None or len(df.columns) <= 1:
            try:
                print("Trying to read as Excel file...")
                df = pd.read_excel(csv_file_path)
                print(f"Excel read successful: {df.shape}, Columns: {list(df.columns)}")
            except Exception as e:
                print(f"Excel read failed: {e}")

        if df is None or len(df.columns) <= 1:
            messagebox.showerror(
                "File Format Error",
                f"Cannot properly parse this file.\n\n"
                + f"Raw first line: {first_line}\n\n"
                + f"The file may not be a properly formatted CSV. "
                + f"Please check that it's saved as a CSV with proper separators.",
            )
            return

        # Clean column names
        df.columns = df.columns.str.strip().str.replace(r'^\ufeff', '', regex=True)

        # Robust column detection with normalization and synonyms
        def normalize_col_name(name):
            if not isinstance(name, str):
                name = str(name)
            name = unicodedata.normalize("NFKC", name)
            name = name.replace("\ufeff", "").strip().lower()
            name = re.sub(r"[()\[\]{}]", " ", name)
            name = re.sub(r"[\s\-_.:/\\]+", " ", name)
            tokens = name.split()
            joined = "".join(tokens)
            joined = re.sub(r"[^a-z0-9]", "", joined)
            return joined

        normalized_map = {col: normalize_col_name(col) for col in df.columns}

        def find_by_candidates(candidates):
            # First pass: exact matches on normalized names
            for original, norm in normalized_map.items():
                if norm in candidates:
                    return original
            # Second pass: substring containment for near-matches
            for original, norm in normalized_map.items():
                if any(c in norm for c in candidates):
                    return original
            return None

        # Candidate synonym sets
        asset_candidates = {
            "assetid", "idasset", "assetno", "assetnumber", "assetcode",
            "assetidentifier", "assetref", "assetreference", "asset"
        }
        country_code_candidates = {
            "countrycode", "countryiso", "countryiso2", "countryiso3",
            "isocode", "iso2", "iso3", "alpha2", "alpha3",
            "countryalpha2", "countryalpha3", "cntrycd", "cntrycode"
        }
        revenue_candidates = {
            "revenue", "totalrevenue", "grossrevenue", "netrevenue",
            "revenueusd", "rev", "salesrevenue", "turnover", "sales"
        }

        # Resolve asset column
        asset_col = find_by_candidates(asset_candidates)

        # Resolve country column, preferring code-like fields
        country_col = find_by_candidates(country_code_candidates)
        if country_col is None:
            # Heuristic: choose a column whose values look like ISO 3166 codes (AA or AAA)
            likely = []
            for original in df.columns:
                series = df[original].astype(str).str.strip().str.upper()
                non_null = series[series != ""].head(200)
                if len(non_null) == 0:
                    continue
                matches = non_null.str.match(r"^[A-Z]{2,3}$", na=False)
                ratio = matches.mean() if len(matches) else 0
                if ratio >= 0.6:
                    likely.append((ratio, original))
            if likely:
                likely.sort(reverse=True)
                country_col = likely[0][1]

        # Resolve revenue column
        revenue_col = find_by_candidates(revenue_candidates)
        if revenue_col is None:
            # Last resort: any column whose normalized name contains 'revenue'
            for original, norm in normalized_map.items():
                if "revenue" in norm:
                    revenue_col = original
                    break

        print(f"Column detection mapping: {normalized_map}")
        print(
            f"Found columns - Asset: '{asset_col}', Country: '{country_col}', Revenue: '{revenue_col}'"
        )

        # Check what we found
        if not asset_col:
            messagebox.showerror(
                "Column Missing",
                f"'Asset ID' column not found.\nAvailable columns: {list(df.columns)}",
            )
            return

        if not country_col:
            messagebox.showerror(
                "Column Missing",
                f"'Country' column not found.\nAvailable columns: {list(df.columns)}",
            )
            return

        if not revenue_col:
            messagebox.showwarning(
                "Revenue Column",
                "No revenue column found. Continuing without revenue calculations.",
            )

        # Normalize data
        df[asset_col] = df[asset_col].astype(str).str.strip().str.upper()
        df[country_col] = df[country_col].astype(str).str.strip().str.upper()

        # Filter
        filtered_df = df[
            (df[asset_col].isin(asset_ids)) & (df[country_col].isin(country_codes))
        ]

        if filtered_df.empty:
            messagebox.showinfo(
                "No Matches",
                f"No matching rows found.\n\nSearched for:\nAsset IDs: {asset_ids}\nCountries: {country_codes}",
            )
        else:
            # Show filtered results in a new scrollable window
            preview_window = tk.Toplevel(root)
            preview_window.title(f"Filtered Results - Found {len(filtered_df)} rows")
            preview_window.geometry("1000x600")

            # Create frame for text widget and scrollbars
            frame = tk.Frame(preview_window)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            # Convert DataFrame to text
            text_data = filtered_df.to_string(index=False)

            text_widget = tk.Text(frame, wrap=tk.NONE, font=("Courier", 10))
            text_widget.insert(tk.END, text_data)
            text_widget.config(state=tk.DISABLED)

            # Add scrollbars
            y_scroll = tk.Scrollbar(
                frame, orient=tk.VERTICAL, command=text_widget.yview
            )
            y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

            x_scroll = tk.Scrollbar(
                frame, orient=tk.HORIZONTAL, command=text_widget.xview
            )
            x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

            text_widget.config(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # Button to save filtered data
            def save_and_close():
                output_path = os.path.splitext(csv_file_path)[0] + "_filtered.csv"
                filtered_df.to_csv(output_path, index=False)

                if revenue_col:
                    numeric_rev = pd.to_numeric(
                        filtered_df[revenue_col]
                        .astype(str)
                        .str.replace(",", "", regex=False)
                        .str.replace("$", "", regex=False)
                        .str.strip(),
                        errors="coerce",
                    )
                    total_revenue = numeric_rev.sum()
                    messagebox.showinfo(
                        "Results Saved",
                        f"Filtered data saved to:\n{output_path}\n\nTotal Revenue: ${total_revenue:,.2f}",
                    )
                else:
                    messagebox.showinfo(
                        "Results Saved", f"Filtered data saved to:\n{output_path}"
                    )
                preview_window.destroy()

            save_button = tk.Button(
                preview_window, text="Save Filtered CSV", command=save_and_close
            )
            save_button.pack(pady=10)

            # Keep the window open
            preview_window.mainloop()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


if __name__ == "__main__":
    search_excel_with_gui()
