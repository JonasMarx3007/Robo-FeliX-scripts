import pandas as pd
from tkinter import Tk, filedialog, messagebox
import os
import string
import sys

TARGET_MASS_UG = 10.0
TARGET_TOTAL_VOL_UL = 30.0
MIN_CONC_LIMIT = 0.33
MAX_CONC_LIMIT = 20.0


def generate_plate_positions(rows=8, cols=12):
    row_labels = list(string.ascii_uppercase[:rows])
    col_labels = list(range(1, cols + 1))
    return [f"{row}{col}" for col in col_labels for row in row_labels]


def alert_and_exit(message, title="Info"):
    root = Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    sys.exit(0)


def error_and_exit(message, title="Error"):
    root = Tk()
    root.withdraw()
    messagebox.showerror(title, message)
    sys.exit(1)


def process_plate_excel():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select an Excel plate file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        alert_and_exit("No file selected.", "Cancelled")

    try:
        df = pd.read_excel(file_path)
        df.rename(columns={df.columns[0]: 'Row'}, inplace=True)

        df_long = df.melt(id_vars='Row', var_name='Column', value_name='Value').dropna().reset_index(drop=True)

        df_long['org_pos'] = df_long['Row'] + df_long['Column'].astype(str)
        df_long.rename(columns={'Value': 'org_con'}, inplace=True)

        df_long['sample_vol'] = TARGET_MASS_UG / df_long['org_con']

        invalid_low_conc = df_long[df_long['org_con'] < MIN_CONC_LIMIT]
        if not invalid_low_conc.empty:
            bad_wells = ", ".join(invalid_low_conc['org_pos'].tolist())
            error_and_exit(
                f"Concentration too low (< {MIN_CONC_LIMIT} µg/µl) in wells:\n{bad_wells}",
                "Lower Concentration Limit Error"
            )

        invalid_high_conc = df_long[df_long['org_con'] > MAX_CONC_LIMIT]
        if not invalid_high_conc.empty:
            bad_wells = ", ".join(invalid_high_conc['org_pos'].tolist())
            error_and_exit(
                f"Concentration too high (> {MAX_CONC_LIMIT} µg/µl) in wells:\n{bad_wells}",
                "Upper Concentration Limit Error"
            )

        df_long['buffer_vol'] = TARGET_TOTAL_VOL_UL - df_long['sample_vol']

        plate_positions = generate_plate_positions()
        if len(df_long) > len(plate_positions):
            error_and_exit(f"Too many samples ({len(df_long)}).", "Plate Overflow")

        df_long['new_pos'] = plate_positions[:len(df_long)]

        output_df = pd.DataFrame({
            'org_pos': df_long['org_pos'],
            'new_pos': df_long['new_pos'],
            'org_con': df_long['org_con'],
            'sample_vol': df_long['sample_vol'].round(1),
            'buffer_vol': df_long['buffer_vol'].round(1)
        })

        output_dir = os.path.join(os.path.dirname(file_path), 'robo_sp3_info')
        os.makedirs(output_dir, exist_ok=True)

        main_file_path = os.path.join(output_dir, 'robo_sp3_info.txt')
        headers = output_df.columns.tolist()
        lines = ['; '.join(headers)]
        for _, row in output_df.iterrows():
            lines.append(
                f"{row['org_pos']}; {row['new_pos']}; {row['org_con']}; {row['sample_vol']}; {row['buffer_vol']}")

        with open(main_file_path, 'w') as f:
            f.write('\n'.join(lines))

        columns_to_export = ['org_pos', 'new_pos', 'sample_vol', 'buffer_vol']

        for col in columns_to_export:
            col_file_path = os.path.join(output_dir, f"{col}.txt")
            content = '\n'.join(output_df[col].astype(str).tolist())
            with open(col_file_path, 'w') as f:
                f.write(content)

        alert_and_exit(f"Success! Output folder created at:\n{output_dir}", "Success")

    except Exception as e:
        error_and_exit(f"An error occurred:\n{str(e)}", "System Error")


if __name__ == "__main__":
    process_plate_excel()