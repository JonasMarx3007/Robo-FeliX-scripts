import tkinter as tk
import sys
import os

ROWS = 8
COLS = 12
WELL_RADIUS = 20
PADDING = 10
WELL_COLORS = {"normal": "white", "selected": "lightblue"}


class PlateGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Plate selector")

        self.canvas_frame = tk.Frame(self)
        self.canvas_frame.pack()

        canvas_width = COLS * (WELL_RADIUS * 2 + PADDING) + PADDING
        canvas_height = ROWS * (WELL_RADIUS * 2 + PADDING) + PADDING
        self.canvas = tk.Canvas(self.canvas_frame, width=canvas_width, height=canvas_height, bg="white")
        self.canvas.pack()

        self.control_frame = tk.Frame(self)
        self.control_frame.pack(pady=5)

        tk.Button(self.control_frame, text="Reset (R)", command=self.reset_selection).pack(side="left", padx=5)
        tk.Button(self.control_frame, text="Export (E)", command=self.export_selection).pack(side="left", padx=5)
        tk.Button(self.control_frame, text="Quit (Q)", command=self.quit_program).pack(side="left", padx=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = tk.Label(self, textvariable=self.status_var, fg="gray")
        self.status_label.pack(pady=2)

        self.wells = {}
        self.selected_wells = set()
        self.drag_start = None
        self.selection_rect = None
        self.dragged = False

        self.draw_plate()
        self.bind_events()

    def draw_plate(self):
        for row in range(ROWS):
            for col in range(COLS):
                x = PADDING + col * (WELL_RADIUS * 2 + PADDING) + WELL_RADIUS
                y = PADDING + row * (WELL_RADIUS * 2 + PADDING) + WELL_RADIUS
                well_id = f"{chr(65 + row)}{col + 1}"

                circle = self.canvas.create_oval(
                    x - WELL_RADIUS, y - WELL_RADIUS, x + WELL_RADIUS, y + WELL_RADIUS,
                    fill=WELL_COLORS["normal"], outline="black", tags=("well", well_id)
                )
                label = self.canvas.create_text(x, y, text=well_id, font=("Arial", 10), tags=("well", well_id))
                self.wells[well_id] = (circle, label)

    def bind_events(self):
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        # Keyboard shortcuts
        self.bind("<r>", lambda e: self.reset_selection())
        self.bind("<e>", lambda e: self.export_selection())
        self.bind("<q>", lambda e: self.quit_program())
        self.focus_set()

    def on_click(self, event):
        self.drag_start = (event.x, event.y)
        self.dragged = False

    def on_drag(self, event):
        self.dragged = True
        if self.selection_rect:
            self.canvas.delete(self.selection_rect)

        x0, y0 = self.drag_start
        x1, y1 = event.x, event.y
        self.selection_rect = self.canvas.create_rectangle(
            x0, y0, x1, y1, outline="red", dash=(2, 2)
        )

    def on_release(self, event):
        if self.selection_rect:
            self.apply_selection_rect()
            self.canvas.delete(self.selection_rect)
            self.selection_rect = None

        else:
            x0, y0 = self.drag_start
            dx = abs(event.x - x0)
            dy = abs(event.y - y0)
            if dx < 10 and dy < 10:
                well = self.get_well_at(event.x, event.y)
                if well:
                    self.toggle_well(well)

        self.drag_start = None
        self.dragged = False

    def apply_selection_rect(self):
        x0, y0 = self.drag_start
        x1 = self.canvas.winfo_pointerx() - self.canvas.winfo_rootx()
        y1 = self.canvas.winfo_pointery() - self.canvas.winfo_rooty()
        x_min, x_max = sorted([x0, x1])
        y_min, y_max = sorted([y0, y1])

        for well_id, (circle_id, _) in self.wells.items():
            coords = self.canvas.coords(circle_id)
            cx = (coords[0] + coords[2]) / 2
            cy = (coords[1] + coords[3]) / 2
            if x_min <= cx <= x_max and y_min <= cy <= y_max:
                self.toggle_well(well_id)

    def get_well_at(self, x, y):
        overlapping = self.canvas.find_overlapping(x, y, x, y)
        for item in overlapping:
            tags = self.canvas.gettags(item)
            if "well" in tags:
                return tags[1]
        return None

    def toggle_well(self, well_id):
        if well_id in self.selected_wells:
            self.deselect_well(well_id)
        else:
            self.select_well(well_id)

    def select_well(self, well_id):
        if well_id not in self.selected_wells:
            self.selected_wells.add(well_id)
            circle, _ = self.wells[well_id]
            self.canvas.itemconfig(circle, fill=WELL_COLORS["selected"])

    def deselect_well(self, well_id):
        if well_id in self.selected_wells:
            self.selected_wells.remove(well_id)
            circle, _ = self.wells[well_id]
            self.canvas.itemconfig(circle, fill=WELL_COLORS["normal"])

    def reset_selection(self):
        for well_id in list(self.selected_wells):
            self.deselect_well(well_id)
        self.status_var.set("Selection cleared")

    def export_selection(self):
        ordered_ids = [
            f"{chr(65 + row)}{col + 1}"
            for col in range(COLS)
            for row in range(ROWS)
        ]
        selected_sorted = [wid for wid in ordered_ids if wid in self.selected_wells]

        file_path = os.path.abspath("samples.txt")
        try:
            with open(file_path, "w") as f:
                for well in selected_sorted:
                    f.write(well + "\n")
            print(f"Exported to: {file_path}")
            self.status_var.set("Export complete")
        except Exception as e:
            self.status_var.set(f"Failed to export: {e}")

    def quit_program(self):
        self.destroy()
        sys.exit(0)


if __name__ == "__main__":
    app = PlateGUI()
    app.mainloop()
