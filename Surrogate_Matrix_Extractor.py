import win32com.client as win32
import os
import sys
import numpy as np
import itertools
import pandas as pd
from pathlib import Path
import customtkinter as ctk
from customtkinter import filedialog
import json

# ==========================================
# 1. STREAM REDIRECTOR (IN-APP TERMINAL)
# ==========================================
class ConsoleRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass

# ==========================================
# 2. ASPEN COM CONNECTION & PATHS
# ==========================================
def connect_to_aspen(aspen_filepath):
    try:
        aspen = win32.Dispatch('Apwn.Document')
        aspen.InitFromFile2(os.path.abspath(aspen_filepath))
        return aspen
    except Exception:
        return None

def scan_aspen_inventory(aspen):
    blocks, streams, components = [], [], []
    try:
        b_node = aspen.Tree.FindNode(r"\Data\Blocks")
        if b_node: blocks = [b.Name for b in b_node.Elements]
            
        s_node = aspen.Tree.FindNode(r"\Data\Streams")
        if s_node: streams = [s.Name for s in s_node.Elements]
        
        comp_paths = [r"\Data\Components\Specifications\Input\ID1", r"\Data\Components\Specifications\Input\COMP-ID"]
        for p in comp_paths:
            c_node = aspen.Tree.FindNode(p)
            if c_node:
                components = [c.Name for c in c_node.Elements]
                break
    except Exception:
        pass
    
    if not components:
        components = ["(Scan Failed - Type manually)"]
        
    return blocks, streams, components

def build_aspen_path(is_input, t_type, target, prop, comp=""):
    # 1. The Manual Override Failsafe
    if "Custom Path" in prop:
        return comp.strip()

    # 2. Input Matrix Routing
    if is_input:
        if t_type == "Block": 
            return rf"\Data\Blocks\{target}\Input\{prop}"
        elif t_type == "Stream":
            if prop == "Component Flow": return rf"\Data\Streams\{target}\Input\FLOW\MIXED\{comp}"
            else: return rf"\Data\Streams\{target}\Input\{prop}"
            
    # 3. Output Extraction Routing
    else:
        if t_type == "Block":
            if prop == "Conversion": return rf"\Data\Blocks\{target}\Output\CONV\{comp}"
            elif prop == "Heat Duty (QCALC)": return rf"\Data\Blocks\{target}\Output\QCALC"
            elif prop == "Vapor Fraction": return rf"\Data\Blocks\{target}\Output\VFRAC_OUT"
            elif prop == "Outlet Temp": return rf"\Data\Blocks\{target}\Output\TEMP_OUT"
            elif prop == "Outlet Pressure": return rf"\Data\Blocks\{target}\Output\PRES_OUT"
            else: return rf"\Data\Blocks\{target}\Output\{prop}"
        elif t_type == "Stream":
            if prop == "Component Mass Flow": return rf"\Data\Streams\{target}\Output\MASSFLOW\MIXED\{comp}"
            elif prop == "Component Mole Flow": return rf"\Data\Streams\{target}\Output\MOLEFLOW\MIXED\{comp}"
            elif prop == "Component Mass Frac": return rf"\Data\Streams\{target}\Output\MASSFRAC\MIXED\{comp}"
            elif prop == "Component Mole Frac": return rf"\Data\Streams\{target}\Output\MOLEFRAC\MIXED\{comp}"
            elif prop == "Total Mass Flow": return rf"\Data\Streams\{target}\Output\RES_MASSFLOW"
            elif prop == "Total Mole Flow": return rf"\Data\Streams\{target}\Output\RES_MOLEFLOW"
            elif prop == "Mass Enthalpy": return rf"\Data\Streams\{target}\Output\MASS-ENTHALPY\MIXED"
            elif prop == "Mass Density": return rf"\Data\Streams\{target}\Output\MASS-RHO\MIXED"
            else: return rf"\Data\Streams\{target}\Output\{prop}"

# --- UPDATED DROPDOWN DICTIONARIES ---
IN_PROPS = {
    "Block": ["TEMP", "PRES", "CTEMP", "HEAT-DUTY", "VFRAC", "NTRAY", "FEED-STAGE", "MOLE-RR", "MASS-RR", "DIST-RATE", "BTM-RATE", "Custom Path (Paste in Component Box)"], 
    "Stream": ["Component Flow", "TOTFLOW", "TEMP", "PRES", "VFRAC", "Custom Path (Paste in Component Box)"]
}

OUT_PROPS = {
    "Block": ["Conversion", "Heat Duty (QCALC)", "Outlet Temp", "Outlet Pressure", "Vapor Fraction", "Custom Path (Paste in Component Box)"], 
    "Stream": ["Component Mass Flow", "Component Mole Flow", "Component Mass Frac", "Component Mole Frac", "Total Mass Flow", "Total Mole Flow", "TEMP", "PRES", "VFRAC", "Mass Enthalpy", "Mass Density", "Custom Path (Paste in Component Box)"]
}
# ==========================================
# 3. GRAPHICAL USER INTERFACE
# ==========================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

IN_PROPS = {
    "Block": ["TEMP", "PRES", "CTEMP", "HEAT-DUTY", "VFRAC", "Custom Path (Paste in Component Box)"], 
    "Stream": ["Component Mass Flow", "TEMP", "PRES", "Custom Path (Paste in Component Box)"]
}
OUT_PROPS = {
    "Block": ["Conversion", "TEMP_OUT", "PRES_OUT", "Heat Duty (QCALC)", "Vapor Fraction", "Custom Path (Paste in Component Box)"], 
    "Stream": ["Component Mass Flow", "Component Mole Flow", "Component Mass Frac", "Component Mole Frac", "Total Mass Flow", "Total Mole Flow", "TEMP", "PRES", "Custom Path (Paste in Component Box)"]
}

class SurrogateMatrixExtractorUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Surrogate Matrix Extractor")
        self.geometry("1250x820")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.aspen_link = None
        self.available_blocks = ["(Load File First)"]
        self.available_streams = ["(Load File First)"]
        self.available_comps = ["(Load File First)"]
        
        self.var_rows = []
        self.constant_rows = []
        self.obj_rows = []
        
        self.abort_flag = False

        self.sidebar = ctk.CTkFrame(self, width=230, corner_radius=0, fg_color="#1a1a2e")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(5, weight=1) 

        self.logo_label = ctk.CTkLabel(self.sidebar, text="📊 SURROGATE\nEXTRACTOR", font=ctk.CTkFont(size=20, weight="bold"), text_color="#e94560")
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 5))
        self.version_label = ctk.CTkLabel(self.sidebar, text="Matrix Grid v3.0", text_color="gray", font=ctk.CTkFont(size=12))
        self.version_label.grid(row=1, column=0, padx=20, pady=(0, 15))

        self.telem_frame = ctk.CTkFrame(self.sidebar, fg_color="gray15", corner_radius=8)
        self.telem_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        
        ctk.CTkLabel(self.telem_frame, text="SYSTEM STATUS", font=ctk.CTkFont(size=11, weight="bold"), text_color="#e94560").pack(anchor="w", padx=10, pady=(8, 2))
        self.telem_status = ctk.CTkLabel(self.telem_frame, text="Idle / Standby", font=ctk.CTkFont(size=12))
        self.telem_status.pack(anchor="w", padx=10, pady=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(self.telem_frame, orientation="horizontal", height=10, progress_color="#e94560")
        self.progress_bar.pack(fill="x", padx=10, pady=(5, 5))
        self.progress_bar.set(0.0)

        self.telem_pct = ctk.CTkLabel(self.telem_frame, text="Extraction Progress: 0%", font=ctk.CTkFont(size=11), text_color="gray")
        self.telem_pct.pack(anchor="w", padx=10, pady=(0, 8))

        self.tabs = ctk.CTkTabview(self, segmented_button_selected_color="#e94560", segmented_button_selected_hover_color="#c81d3d")
        self.tabs.grid(row=0, column=1, sticky="nsew", padx=20, pady=(10, 0))

        self.tab_engine = self.tabs.add("1. Engine & File")
        self.tab_vars = self.tabs.add("2. Grid Matrix (Inputs)")
        self.tab_consts_setup = self.tabs.add("3. Constants")
        self.tab_objs = self.tabs.add("4. Records (Outputs)")

        self.setup_engine_tab()
        self.setup_vars_tab()
        self.setup_constants_tab()
        self.setup_objs_tab()

        self.bottom_frame = ctk.CTkFrame(self, height=180, corner_radius=0, fg_color="gray10")
        self.bottom_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        term_header = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", height=25)
        term_header.pack(fill="x", padx=15, pady=(5, 0))
        ctk.CTkLabel(term_header, text="MATRIX EXTRACTOR CONSOLE LOG", font=ctk.CTkFont(size=11, weight="bold"), text_color="#e94560").pack(side="left")

        self.console_box = ctk.CTkTextbox(self.bottom_frame, height=140, font=ctk.CTkFont(family="Consolas", size=11), state="disabled", fg_color="black", text_color="#ffcc00")
        self.console_box.pack(fill="both", expand=True, padx=15, pady=(2, 10))

        sys.stdout = ConsoleRedirector(self.console_box)
        print("[CONSOLE] Surrogate Matrix Extractor Initialized.")

    def delete_row(self, row_data, row_list):
        row_data["frame"].destroy()
        row_list.remove(row_data)

    def setup_engine_tab(self):
        self.file_label = ctk.CTkLabel(self.tab_engine, text="Aspen Project File:", font=ctk.CTkFont(weight="bold"))
        self.file_label.pack(anchor="w", padx=20, pady=(15, 5))
        
        frame1 = ctk.CTkFrame(self.tab_engine, fg_color="transparent")
        frame1.pack(fill="x", padx=20)
        
        self.file_entry = ctk.CTkEntry(frame1, placeholder_text="Select your .apwz file...")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(frame1, text="Browse", width=80, fg_color="#4a4e69", hover_color="#22223b", command=self.browse_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(frame1, text="Connect & Scan", width=130, fg_color="#b87333", hover_color="#8c5827", command=self.scan_file).pack(side="left")

        preset_frame = ctk.CTkFrame(self.tab_engine, fg_color="gray20", corner_radius=8)
        preset_frame.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkButton(preset_frame, text="Save Config", width=90, fg_color="#27ae60", hover_color="#1e8449", command=self.save_config).pack(side="left", padx=10, pady=12)
        ctk.CTkButton(preset_frame, text="Load Config", width=90, fg_color="#2980b9", hover_color="#1f618d", command=self.load_config).pack(side="left", padx=(0, 10))
        ctk.CTkButton(preset_frame, text="Test Single Run", width=140, fg_color="#8e44ad", hover_color="#5b2c6f", command=self.test_paths).pack(side="left", padx=(0, 10))

        self.run_btn = ctk.CTkButton(
            self.tab_engine, text="LAUNCH MATRIX EXTRACTION (EXCEL EXPORT)", font=ctk.CTkFont(size=16, weight="bold"), 
            fg_color="#e94560", hover_color="#a23043", height=45, command=self.toggle_engine_state
        )
        self.run_btn.pack(fill="x", padx=20, pady=25)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Aspen Plus Workspaces", "*.apwz")])
        if filename:
            self.file_entry.delete(0, 'end'); self.file_entry.insert(0, filename)
            print(f"[ASPEN] Selected workspace: {filename}")

    def scan_file(self):
        filepath = self.file_entry.get()
        if not filepath or not os.path.exists(filepath): return
        print("[ASPEN] Initializing COM link and cataloging plant topology...")
        self.telem_status.configure(text="Scanning Topology...", text_color="#f39c12"); self.update()
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if self.aspen_link:
            blocks, streams, comps = scan_aspen_inventory(self.aspen_link)
            self.available_blocks = blocks if blocks else ["None Found"]
            self.available_streams = streams if streams else ["None Found"]
            self.available_comps = comps if comps else ["None Found"]
            self.refresh_dropdowns()
            self.telem_status.configure(text="Connected / Ready", text_color="#1d865c")
            print(f"[ASPEN] Scan successful: Identified {len(blocks)} Blocks, {len(streams)} Streams, and {len(comps)} Components.")

    def refresh_dropdowns(self):
        for row in self.var_rows + self.constant_rows: 
            row['target'].configure(values=self.available_blocks if row['type'].get() == "Block" else self.available_streams)
            row['comp'].configure(values=self.available_comps)
        for row in self.obj_rows:
            row['target'].configure(values=self.available_blocks if row['type'].get() == "Block" else self.available_streams)
            row['comp'].configure(values=self.available_comps)

    def _create_dynamic_row(self, parent, is_input, vals):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", pady=2)
        
        props_dict = IN_PROPS if is_input else OUT_PROPS
        
        e_type = ctk.CTkOptionMenu(row, width=90, values=["Block", "Stream"])
        e_target = ctk.CTkComboBox(row, width=120, values=self.available_blocks)
        e_prop = ctk.CTkComboBox(row, width=220, values=props_dict["Block"])
        e_comp = ctk.CTkComboBox(row, width=150, values=self.available_comps)
        
        def update_dropdowns(choice, prop_val=None):
            if choice == "Block":
                e_target.configure(values=self.available_blocks)
                if not e_target.get() in self.available_blocks: e_target.set(self.available_blocks[0] if self.available_blocks else "")
                e_prop.configure(values=props_dict["Block"])
                e_prop.set(prop_val if prop_val in props_dict["Block"] else props_dict["Block"][0])
                e_comp.set(""); e_comp.configure(state="disabled")
            else:
                e_target.configure(values=self.available_streams)
                if not e_target.get() in self.available_streams: e_target.set(self.available_streams[0] if self.available_streams else "")
                e_prop.configure(values=props_dict["Stream"])
                e_prop.set(prop_val if prop_val in props_dict["Stream"] else props_dict["Stream"][0])
            check_comp_state(e_prop.get())

        def check_comp_state(choice):
            if "Component" in choice or "Conversion" in choice:
                e_comp.configure(state="normal")
                if not e_comp.get() or e_comp.get() == "": e_comp.set(self.available_comps[0] if self.available_comps else "")
            elif "Custom Path" in choice:
                e_comp.configure(state="normal")
                e_comp.set(vals.get("comp", r"\Data\Blocks\..."))
            else:
                e_comp.set(""); e_comp.configure(state="disabled")

        e_type.configure(command=lambda c: update_dropdowns(c))
        e_prop.configure(command=check_comp_state)
        
        e_type.set(vals.get("type", "Block"))
        update_dropdowns(e_type.get(), vals.get("prop", ""))
        e_target.set(vals.get("target", ""))
        e_comp.set(vals.get("comp", ""))
        
        e_type.pack(side="left", padx=5); e_target.pack(side="left", padx=5)
        e_prop.pack(side="left", padx=5); e_comp.pack(side="left", padx=5)
        
        row_data = {"frame": row, "type": e_type, "target": e_target, "prop": e_prop, "comp": e_comp}
        return row, row_data

    def setup_vars_tab(self):
        btn_frame = ctk.CTkFrame(self.tab_vars, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkButton(btn_frame, text="+ Add Grid Variable", fg_color="#e94560", hover_color="#c81d3d", command=lambda: self.add_var_row({})).pack(side="left", padx=10)
        
        h_frame = ctk.CTkFrame(self.tab_vars, fg_color="transparent", height=20)
        h_frame.pack(fill="x", padx=20, pady=(15, 0))
        ctk.CTkLabel(h_frame, text="Type", width=90, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Target", width=120, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Property", width=220, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Component / Custom Path", width=150, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        
        # BUG FIX: text_color is an argument of CTkLabel, not CTkFont
        ctk.CTkLabel(h_frame, text="Min", width=70, anchor="w", font=ctk.CTkFont(weight="bold"), text_color="#2ecc71").pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Max", width=70, anchor="w", font=ctk.CTkFont(weight="bold"), text_color="#e74c3c").pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Step", width=70, anchor="w", font=ctk.CTkFont(weight="bold"), text_color="#3498db").pack(side="left", padx=5)

        self.vars_scroll = ctk.CTkScrollableFrame(self.tab_vars, fg_color="transparent")
        self.vars_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_var_row(self, vals):
        row, row_data = self._create_dynamic_row(self.vars_scroll, True, vals)
        e_min = ctk.CTkEntry(row, width=70, placeholder_text="Min"); e_min.insert(0, vals.get("min", "")); e_min.pack(side="left", padx=5)
        e_max = ctk.CTkEntry(row, width=70, placeholder_text="Max"); e_max.insert(0, vals.get("max", "")); e_max.pack(side="left", padx=5)
        e_step = ctk.CTkEntry(row, width=70, placeholder_text="Step"); e_step.insert(0, vals.get("step", "")); e_step.pack(side="left", padx=5)
        row_data.update({"min": e_min, "max": e_max, "step": e_step})
        self.var_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.var_rows)).pack(side="left", padx=5)

    def setup_constants_tab(self):
        ctk.CTkButton(self.tab_consts_setup, text="+ Add Constant", command=lambda: self.add_constant_row({})).pack(anchor="w", padx=20, pady=10)
        self.constants_scroll = ctk.CTkScrollableFrame(self.tab_consts_setup, fg_color="transparent")
        self.constants_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_constant_row(self, vals):
        row, row_data = self._create_dynamic_row(self.constants_scroll, True, vals)
        e_val = ctk.CTkEntry(row, width=120, placeholder_text="Static Value"); e_val.insert(0, vals.get("val", "")); e_val.pack(side="left", padx=5)
        row_data.update({"val": e_val})
        self.constant_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.constant_rows)).pack(side="left", padx=5)

    def setup_objs_tab(self):
        btn_frame = ctk.CTkFrame(self.tab_objs, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkButton(btn_frame, text="+ Add Output Record", command=lambda: self.add_obj_row({})).pack(side="left", padx=10)
        
        h_frame = ctk.CTkFrame(self.tab_objs, fg_color="transparent", height=20)
        h_frame.pack(fill="x", padx=20, pady=(15, 0))
        ctk.CTkLabel(h_frame, text="Type", width=90, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Target", width=120, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Property", width=220, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(h_frame, text="Component / Custom Path", width=150, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)

        self.objs_scroll = ctk.CTkScrollableFrame(self.tab_objs, fg_color="transparent")
        self.objs_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_obj_row(self, vals):
        row, row_data = self._create_dynamic_row(self.objs_scroll, False, vals)
        self.obj_rows.append(row_data)
        ctk.CTkButton(row, text="X", width=30, fg_color="#c0392b", hover_color="#922b21", command=lambda: self.delete_row(row_data, self.obj_rows)).pack(side="left", padx=5)

    def test_paths(self):
        print("\n" + "="*60 + "\n[DIAGNOSTIC] Verifying Paths and Running Single Test...")
        filepath = self.file_entry.get()
        if not filepath:
            print("[ERROR] Please select your Aspen file first."); return
        if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
        if not self.aspen_link:
            print("[ERROR] Could not connect."); return

        try:
            print("[1/2] Checking Inputs & Applying Minimum Values...")
            for row in self.var_rows:
                t_type, target, prop, comp = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get()
                m = row['min'].get().strip()
                if not target or not m or "Load File" in target: continue
                path_str = build_aspen_path(True, t_type, target, prop, comp)
                node = self.aspen_link.Tree.FindNode(path_str)
                if node:
                    node.Value = float(m)
                    print(f"  [OK] Input {target}_{prop} set to Min ({m})")
                else: print(f"  [FAIL] Could not locate input path: {path_str}")

            print("\n[2/2] Running Aspen and Fetching Outputs...")
            self.aspen_link.Engine.Run2()
            
            error_node = self.aspen_link.Tree.FindNode(r"\Data\Results Summary\Run-Status\NERROR")
            if error_node and error_node.Value is not None and int(error_node.Value) > 0:
                print(f"  [WARNING] Aspen returned {int(error_node.Value)} errors during calculation.")

            for row in self.obj_rows:
                t_type, target, prop, comp = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get()
                if not target or "Load File" in target: continue
                path_str = build_aspen_path(False, t_type, target, prop, comp)
                node = self.aspen_link.Tree.FindNode(path_str)
                if node and node.Value is not None:
                    print(f"  [OK] Output {target}_{prop} = {node.Value:<15.4f}")
                else: print(f"  [FAIL] Could not locate output path: {path_str}")
                
            print("[SUCCESS] Test run complete.")
        except Exception as e:
            print(f"[ERROR] Test failed: {str(e)}")
        print("="*60 + "\n")

    def save_config(self):
        data = {"vars": [], "constants": [], "objs": []}
        for r in self.var_rows: data["vars"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "min": r['min'].get(), "max": r['max'].get(), "step": r['step'].get()})
        for r in self.constant_rows: data["constants"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get(), "val": r['val'].get()})
        for r in self.obj_rows: data["objs"].append({"type": r['type'].get(), "target": r['target'].get(), "prop": r['prop'].get(), "comp": r['comp'].get()})
        
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'w') as f: json.dump(data, f, indent=4)
            print(f"[SYSTEM] Config saved: {filepath}")

    def load_config(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filepath:
            with open(filepath, 'r') as f: data = json.load(f)
            for r in list(self.var_rows): self.delete_row(r, self.var_rows)
            for r in list(self.constant_rows): self.delete_row(r, self.constant_rows)
            for r in list(self.obj_rows): self.delete_row(r, self.obj_rows)
            for v in data.get("vars", []): self.add_var_row(v)
            for c in data.get("constants", []): self.add_constant_row(c)
            for o in data.get("objs", []): self.add_obj_row(o)
            print(f"[SYSTEM] Config loaded: {filepath}")

    def toggle_engine_state(self):
        if self.run_btn.cget("text") == "LAUNCH MATRIX EXTRACTION (EXCEL EXPORT)":
            self.abort_flag = False
            self.run_btn.configure(text="🛑 ABORT OPERATION (SAFE STOP)", fg_color="#c0392b", hover_color="#922b21")
            self.launch_harvester()
        else:
            print("[SYSTEM] User initiated Emergency Stop..."); self.abort_flag = True

    def reset_run_btn(self):
        self.run_btn.configure(text="LAUNCH MATRIX EXTRACTION (EXCEL EXPORT)", fg_color="#e94560", hover_color="#a23043")

    def launch_harvester(self):
        filepath = self.file_entry.get()
        if not filepath: 
            print("[ERROR] Flowsheet not loaded."); self.reset_run_btn(); return
            
        config_inputs = []
        grid_axes = []
        
        try:
            for i, row in enumerate(self.var_rows):
                t_type, target, prop, comp = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get()
                m, mx, stp = row['min'].get().strip(), row['max'].get().strip(), row['step'].get().strip()
                
                if not target or not m or not mx or not stp or "Load File" in target: continue
                path_str = build_aspen_path(True, t_type, target, prop, comp)
                
                clean_prop = "Custom_Path" if "Custom" in prop else prop.replace(chr(92), "_").replace(" ", "")
                name_str = f"{target}_{comp if comp else clean_prop}"
                
                vmin, vmax, vstep = float(m), float(mx), float(stp)
                if vmin > vmax: vmin, vmax = vmax, vmin
                if vstep <= 0: 
                    print(f"[ERROR] Step size for {name_str} must be > 0.")
                    self.reset_run_btn(); return
                
                axis_values = np.arange(vmin, vmax + (vstep / 2), vstep)
                grid_axes.append(axis_values)
                config_inputs.append({"name": name_str, "path": path_str})
                
            config_outputs = []
            for row in self.obj_rows:
                t_type, target, prop, comp = row['type'].get(), row['target'].get(), row['prop'].get(), row['comp'].get()
                if not target or "Load File" in target: continue
                path_str = build_aspen_path(False, t_type, target, prop, comp)
                
                clean_prop = "Custom_Path" if "Custom" in prop else prop.replace(chr(92), "_").replace(" ", "")
                name_str = f"{target}_{comp if comp else clean_prop}"
                config_outputs.append({"name": name_str, "path": path_str})
                
            if len(config_inputs) < 1 or len(config_outputs) < 1:
                print("[ERROR] Need >=1 Variable and >=1 Output to record."); self.reset_run_btn(); return
        except ValueError:
            print("[ERROR] Type parsing failure: Check Min/Max/Step formats."); self.reset_run_btn(); return

        all_combinations = list(itertools.product(*grid_axes))
        total_runs = len(all_combinations)
        
        print(f"\n[ENGINE] Matrix compiled. Total permutations to simulate: {total_runs}")
        estimated_seconds = total_runs * 2
        print(f"[INFO] Estimated extraction time: {estimated_seconds/60:.1f} minutes.")
        
        if total_runs > 5000:
            print("[WARNING] Massive dataset requested. Application may appear frozen during data extraction.")

        self.telem_status.configure(text="Extracting Data...", text_color="#2ecc71"); self.progress_bar.set(0.0); self.update()

        try:
            if not self.aspen_link: self.aspen_link = connect_to_aspen(filepath)
            if not self.aspen_link: print("[ERROR] Aspen COM link failed."); return

            harvested_data = []

            for run_idx, combo in enumerate(all_combinations):
                if self.abort_flag:
                    print("[SYSTEM] Extraction Aborted. Exporting salvaged data...")
                    break
                    
                data_row = {}
                for i, val in enumerate(combo):
                    node = self.aspen_link.Tree.FindNode(config_inputs[i]["path"])
                    if node: node.Value = val
                    data_row[config_inputs[i]["name"]] = round(val, 4)
                    
                self.aspen_link.Engine.Run2()
                
                error_node = self.aspen_link.Tree.FindNode(r"\Data\Results Summary\Run-Status\NERROR")
                if error_node and error_node.Value is not None and int(error_node.Value) > 0:
                    for out in config_outputs: data_row[out["name"]] = "ERROR"
                else:
                    for out in config_outputs:
                        node = self.aspen_link.Tree.FindNode(out["path"])
                        if node and node.Value is not None:
                            data_row[out["name"]] = round(float(node.Value), 4)
                        else:
                            data_row[out["name"]] = "N/A"
                            
                harvested_data.append(data_row)
                
                pct = (run_idx + 1) / total_runs
                self.telem_pct.configure(text=f"Extraction Progress: {int(pct * 100)}% ({run_idx+1}/{total_runs})")
                self.progress_bar.set(pct); self.update()

            if harvested_data:
                df = pd.DataFrame(harvested_data)
                
                desktop_path = Path.home() / "Desktop"
                if not desktop_path.exists(): desktop_path = Path.home() / "OneDrive" / "Desktop"
                
                excel_path = desktop_path / "Surrogate_Training_Matrix.xlsx"
                csv_path = desktop_path / "Surrogate_Training_Matrix.csv"
                
                try:
                    df.to_excel(excel_path, index=False)
                    print(f"\n[SUCCESS] Excel Matrix exported to: {excel_path}")
                except ImportError:
                    df.to_csv(csv_path, index=False)
                    print(f"\n[SUCCESS] (openpyxl not found). CSV Matrix exported to: {csv_path}")

            self.telem_status.configure(text="Extraction Complete", text_color="#1d865c")
            
        except Exception as e:
            print(f"[FATAL] Engine failed: {str(e)}")
        finally:
            self.reset_run_btn()

if __name__ == "__main__":
    app = SurrogateMatrixExtractorUI()
    app.mainloop()