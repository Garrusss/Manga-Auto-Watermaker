import tkinter
import tkinter.messagebox
import customtkinter as ctk
from tkinter import filedialog
import threading
import os
import sys
import json
import shutil
import zipfile
from PIL import Image, ImageStat, UnidentifiedImageError
import traceback
import tempfile
import subprocess
import sys
import os


if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    print("Attempting to redirect stdout/stderr...")
    try:
        sys.stdout = open(os.devnull, 'w')
        sys.stderr = open(os.devnull, 'w')
        print("stdout/stderr redirected.")
    except Exception as e:
        print(f"Warning: Failed to redirect stdout/stderr: {e}", file=sys.__stderr__)

Image.MAX_IMAGE_PIXELS = None

CONFIG_FILENAME = ".AutoWatermarkerConfig.json"
OUTPUT_SUFFIX = "_Processed"
SHORT_IMAGE_SEARCH_STEP = 1000
EXTENSIONS_PSD_PSB = ('.psd', '.psb')
EXTENSIONS_PNG_JPG = ('.png', '.jpg', '.jpeg')
ALLOWED_WATERMARK_EXTENSIONS = ('.png', '.jpg', '.jpeg')
DEFAULT_CONFIG = {
    "main_folder": "", "watermark_file": "", "frequency": "10000",
    "search_step": "300", "threshold": "25", "max_steps": "10",
    "create_zip": False, "magick_path": "", "process_type": "png"
}
DEFAULT_IMAGEMAGICK_COMMAND = "magick"

def is_int(value):
    try: int(value); return True
    except ValueError: return False

def get_config_path():
    home_dir = os.path.expanduser("~")
    preferred_dir = os.path.join(home_dir, ".config", "AutoWatermarker")
    fallback_dir = os.path.join(home_dir, "AutoWatermarker")
    final_dir = None
    try: os.makedirs(preferred_dir, exist_ok=True); final_dir = preferred_dir
    except OSError:
        print(f"Warning: Could not create config directory '{preferred_dir}'. Trying fallback.")
        try: os.makedirs(fallback_dir, exist_ok=True); final_dir = fallback_dir; print(f"Using fallback config directory: {final_dir}")
        except OSError as e: print(f"ERROR: Could not create any config directory. Saving in home directory. Error: {e}"); final_dir = home_dir
    return os.path.join(final_dir, CONFIG_FILENAME)

def convert_to_temp_png(original_path, temp_dir, magick_executable_path, status_callback):
    original_filename = os.path.basename(original_path)
    base_name = os.path.splitext(original_filename)[0]
    temp_png_filename = base_name + "_temp.png"
    temp_png_path = os.path.join(temp_dir, temp_png_filename)
    original_format = os.path.splitext(original_filename)[1].lower()
    conversion_success = False

    status_callback(f"  Converting/copying: {original_filename} -> PNG...")

    try:
        if original_format == ".png":
            shutil.copy2(original_path, temp_png_path)
            conversion_success = True
        elif original_format in ('.psd', '.psb'):
            status_callback(f"   (Using ImageMagick: {magick_executable_path})")
            input_spec = f"{original_path}[0]"
            command_list = [magick_executable_path, input_spec, temp_png_path]
            try:
                run_result = subprocess.run(command_list, check=True, capture_output=True, text=True, encoding='utf-8', errors='replace', startupinfo=None)
                conversion_success = True
            except FileNotFoundError:
                 status_callback(f"  ! ERROR: ImageMagick path '{magick_executable_path}' not found.")
                 conversion_success = False
            except subprocess.CalledProcessError as process_error:
                 status_callback(f"  ! ImageMagick error converting {original_filename}")
                 print(f"--- IMAGEMAGICK ERROR for {original_path} ---\nCommand: {' '.join(process_error.cmd)}\nReturn Code: {process_error.returncode}\nStderr:\n{process_error.stderr}\nStdout:\n{process_error.stdout}\n--- END IMAGEMAGICK ERROR ---")
                 conversion_success = False
            except Exception as subprocess_error:
                 status_callback(f"  ! Unexpected subprocess error: {type(subprocess_error).__name__}")
                 print(f"--- SUBPROCESS ERROR for {original_path} ---\n{traceback.format_exc()}\n--- END SUBPROCESS ERROR ---")
                 conversion_success = False
        elif original_format in ('.jpg', '.jpeg'):
            with Image.open(original_path) as img, img.convert("RGBA") as rgba_img:
                rgba_img.save(temp_png_path, "PNG")
            conversion_success = True
        else:
             status_callback(f"  ! Unsupported format: {original_filename}")

        if conversion_success and os.path.exists(temp_png_path) and os.path.getsize(temp_png_path) > 0:
            return temp_png_path
        else:
            if conversion_success: status_callback(f"  ! Error: Conversion of {original_filename} created an empty file.")
            if os.path.exists(temp_png_path):
                 try: os.remove(temp_png_path)
                 except OSError as remove_error: print(f"Warning: Could not remove temp file: {temp_png_path}. Error: {remove_error}")
            return None

    except Exception as top_level_error:
        status_callback(f"  ! Critical error during conversion stage for {original_filename}: {type(top_level_error).__name__}")
        print(f"--- TOP LEVEL CONVERSION STAGE ERROR for {original_path} ---\n{traceback.format_exc()}\n--- END ERROR ---")
        if 'temp_png_path' in locals() and os.path.exists(temp_png_path):
             try: os.remove(temp_png_path)
             except OSError: pass
        return None

def check_area_uniformity(image, x, y, width, height, threshold):
    try:
        with image.crop((x, y, x + width, y + height)) as area:
            area_gray = area.convert('L'); stat = ImageStat.Stat(area_gray)
            min_val, max_val = stat.extrema[0]; difference = max_val - min_val
            return difference <= threshold
    except Exception as e: print(f"  Warning: Error checking uniformity at ({x},{y}): {e}"); return False

def search_and_place_watermark(main_img, watermark_img, config, start_y, max_search_y):
    main_w, main_h = main_img.size; wm_w, wm_h = watermark_img.size
    placement_x = main_w - wm_w; placement_y = -1; found_spot = False
    if start_y < 0: start_y = 0
    if placement_x < 0: return -1
    if start_y + wm_h <= main_h:
        if check_area_uniformity(main_img, placement_x, start_y, wm_w, wm_h, config['threshold']):
            placement_y = start_y; found_spot = True
        else:
            search_y = start_y + config['search_step']; steps_taken = 0
            effective_max_y = min(max_search_y, main_h - wm_h)
            while search_y <= effective_max_y and steps_taken < config['max_steps']:
                if check_area_uniformity(main_img, placement_x, search_y, wm_w, wm_h, config['threshold']):
                    placement_y = search_y; found_spot = True; break
                search_y += config['search_step']; steps_taken += 1
    if found_spot:
        try: main_img.paste(watermark_img, (placement_x, placement_y), watermark_img); return placement_y
        except Exception as e: print(f"  ERROR pasting watermark at Y={placement_y}: {e}"); return -1
    else: return -1

def add_watermarks_to_image(input_png_path, watermark_path, output_final_path, config, status_callback):
    try:
        with Image.open(input_png_path).convert("RGBA") as main_img, \
             Image.open(watermark_path).convert("RGBA") as watermark_img:
            main_w, main_h = main_img.size; wm_w, wm_h = watermark_img.size
            watermarks_added_count = 0
            if wm_w > main_w or wm_h > main_h: status_callback(f"  - Watermark larger than image.")
            elif main_h < config['frequency']:
                placement_x = main_w - wm_w; search_y = 0; found_spot_short = False
                while search_y + wm_h <= main_h:
                    if check_area_uniformity(main_img, placement_x, search_y, wm_w, wm_h, config['threshold']):
                        try: main_img.paste(watermark_img, (placement_x, search_y), watermark_img); watermarks_added_count = 1; found_spot_short = True; status_callback(f"  + Watermark (Y={search_y})"); break
                        except Exception as e: status_callback(f"  ! Error applying watermark (Y={search_y}): {e}"); break
                    search_y += SHORT_IMAGE_SEARCH_STEP
                if not found_spot_short and watermarks_added_count == 0: status_callback(f"  - Spot not found (short image).")
            else:
                current_y_target = config['frequency']
                while current_y_target < main_h:
                    if current_y_target + wm_h <= main_h:
                        max_y_for_interval_search = current_y_target + config['frequency']
                        placement_y = search_and_place_watermark(main_img, watermark_img, config, current_y_target, max_y_for_interval_search)
                        if placement_y != -1: watermarks_added_count += 1; status_callback(f"  + Watermark (Y={placement_y})")
                    else: break
                    current_y_target += config['frequency']

            output_dir = os.path.dirname(output_final_path)
            if output_dir:
                try: os.makedirs(output_dir, exist_ok=True)
                except OSError as e: status_callback(f"  ! Error creating folder '{output_dir}': {e}"); return False
            if watermarks_added_count > 0:
                try: main_img.save(output_final_path, "PNG"); return True
                except Exception as e: status_callback(f"  ! Error saving result: {e}"); print(f"--- ERROR SAVING RESULT for {output_final_path} ---\n{traceback.format_exc()}\n--- END ERROR ---"); return False
            else:
                try: shutil.copy2(input_png_path, output_final_path); return True
                except Exception as e: status_callback(f"  ! Error copying PNG: {e}"); print(f"--- ERROR COPYING temp PNG {input_png_path} to {output_final_path} ---\n{traceback.format_exc()}\n--- END ERROR ---"); return False
    except FileNotFoundError: status_callback(f"  ! Error: PNG '{os.path.basename(input_png_path)}' or watermark not found."); return False
    except UnidentifiedImageError: status_callback(f"  ! Error: Could not identify PNG format '{os.path.basename(input_png_path)}'."); return False
    except Exception as e: status_callback(f"  ! Error processing PNG '{os.path.basename(input_png_path)}': {type(e).__name__}"); print(f"--- WATERMARKING ERROR for {input_png_path} ---\n{traceback.format_exc()}\n--- END ERROR ---"); return False

class WatermarkerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto Watermarker V2.5 (IM Path Config)")
        self.geometry("750x800")
        ctk.set_appearance_mode("Dark"); ctk.set_default_color_theme("blue")
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(6, weight=1)
        self._initialize_state()
        self.load_settings()
        self._create_widgets()
        self._validate_loaded_magick_path()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _initialize_state(self):
        self.main_folder = tkinter.StringVar(); self.watermark_file = tkinter.StringVar()
        self.watermark_dims = tkinter.StringVar(value="Size: (not selected)")
        self.frequency = tkinter.StringVar(); self.search_step = tkinter.StringVar()
        self.threshold = tkinter.StringVar(); self.max_steps = tkinter.StringVar()
        self.create_zip = tkinter.BooleanVar()
        self.magick_path_var = tkinter.StringVar(); self.verified_magick_path = None
        self.process_type = tkinter.StringVar(value="png")

    def _create_widgets(self):
        current_row = 0
        self.select_frame = ctk.CTkFrame(self); self.select_frame.grid(row=current_row, column=0, padx=20, pady=(20, 10), sticky="ew"); self.select_frame.grid_columnconfigure(1, weight=1)
        self.main_folder_btn = ctk.CTkButton(self.select_frame, text="Main Folder", command=self.select_main_folder); self.main_folder_btn.grid(row=0, column=0, padx=(20, 10), pady=10)
        self.main_folder_label = ctk.CTkLabel(self.select_frame, textvariable=self.main_folder, anchor="w", text="..."); self.main_folder_label.grid(row=0, column=1, padx=(0, 20), pady=10, sticky="ew")
        self.watermark_btn = ctk.CTkButton(self.select_frame, text="Watermark File", command=self.select_watermark_file); self.watermark_btn.grid(row=1, column=0, padx=(20, 10), pady=10)
        watermark_info_frame = ctk.CTkFrame(self.select_frame, fg_color="transparent"); watermark_info_frame.grid(row=1, column=1, padx=(0, 20), pady=10, sticky="ew"); watermark_info_frame.grid_columnconfigure(0, weight=1)
        self.watermark_label = ctk.CTkLabel(watermark_info_frame, textvariable=self.watermark_file, anchor="w", text="..."); self.watermark_label.grid(row=0, column=0, sticky="ew")
        self.watermark_dims_label = ctk.CTkLabel(watermark_info_frame, textvariable=self.watermark_dims, anchor="e", text_color="gray"); self.watermark_dims_label.grid(row=0, column=1, padx=(10, 0), sticky="e")
        current_row += 1

        self.settings_frame = ctk.CTkFrame(self); self.settings_frame.grid(row=current_row, column=0, padx=20, pady=10, sticky="ew"); self.settings_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="settings_cols")
        ctk.CTkLabel(self.settings_frame, text="Frequency (px):").grid(row=0, column=0, padx=(20, 5), pady=10, sticky="w"); self.freq_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.frequency, width=80); self.freq_entry.grid(row=0, column=1, padx=(0, 15), pady=10, sticky="w")
        ctk.CTkLabel(self.settings_frame, text="Search Step (px):").grid(row=0, column=2, padx=(5, 5), pady=10, sticky="w"); self.step_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.search_step, width=80); self.step_entry.grid(row=0, column=3, padx=(0, 15), pady=10, sticky="w")
        ctk.CTkLabel(self.settings_frame, text="Uniformity Thresh.:").grid(row=1, column=0, padx=(20, 5), pady=10, sticky="w"); self.thresh_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.threshold, width=80); self.thresh_entry.grid(row=1, column=1, padx=(0, 15), pady=10, sticky="w")
        ctk.CTkLabel(self.settings_frame, text="Max Steps:").grid(row=1, column=2, padx=(5, 5), pady=10, sticky="w"); self.max_steps_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.max_steps, width=80); self.max_steps_entry.grid(row=1, column=3, padx=(0, 15), pady=10, sticky="w")
        current_row += 1

        self.magick_frame = ctk.CTkFrame(self); self.magick_frame.grid(row=current_row, column=0, padx=20, pady=10, sticky="ew"); self.magick_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.magick_frame, text="Path to magick.exe:").grid(row=0, column=0, padx=(20, 10), pady=10, sticky="w")
        self.magick_path_entry = ctk.CTkEntry(self.magick_frame, textvariable=self.magick_path_var); self.magick_path_entry.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="ew")
        self.magick_browse_btn = ctk.CTkButton(self.magick_frame, text="Browse...", width=70, command=self._browse_magick_path); self.magick_browse_btn.grid(row=0, column=2, padx=(0, 10), pady=10)
        self.magick_check_btn = ctk.CTkButton(self.magick_frame, text="Check and Save", command=self._check_and_save_magick_path); self.magick_check_btn.grid(row=0, column=3, padx=(0, 20), pady=10)
        current_row += 1

        self.file_type_frame = ctk.CTkFrame(self); self.file_type_frame.grid(row=current_row, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(self.file_type_frame, text="File Type to Process:").pack(side="left", padx=(20, 10), pady=10)
        self.png_radio_button = ctk.CTkRadioButton(self.file_type_frame, text="PNG / JPEG", variable=self.process_type, value="png"); self.png_radio_button.pack(side="left", padx=10, pady=10)
        self.psd_radio_button = ctk.CTkRadioButton(self.file_type_frame, text="PSD / PSB (requires ImageMagick)", variable=self.process_type, value="psd", state="disabled"); self.psd_radio_button.pack(side="left", padx=10, pady=10)
        current_row += 1

        self.action_frame = ctk.CTkFrame(self); self.action_frame.grid(row=current_row, column=0, padx=20, pady=10, sticky="ew"); self.action_frame.grid_columnconfigure(0, weight=1)
        self.zip_checkbox = ctk.CTkCheckBox(self.action_frame, text="Create ZIP archives for each chapter", variable=self.create_zip); self.zip_checkbox.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="w")
        self.start_button = ctk.CTkButton(self.action_frame, text="Start Processing", command=self.start_processing_thread, height=35, font=("Segoe UI", 14, "bold")); self.start_button.grid(row=1, column=0, padx=20, pady=(5, 10), sticky="ew")
        self.progress_bar = ctk.CTkProgressBar(self.action_frame, orientation="horizontal", height=15); self.progress_bar.set(0); self.progress_bar.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")
        current_row += 1

        self.copyright_label = ctk.CTkLabel(self, text="Created by the master of the Garrus team - Vlad", text_color="gray", font=("Segoe UI", 9)); self.copyright_label.grid(row=current_row, column=0, padx=20, pady=(5, 10), sticky="sw")
        current_row += 1

        self.status_textbox = ctk.CTkTextbox(self, wrap="word", height=200, state="disabled", font=("Consolas", 11)); self.status_textbox.grid(row=current_row, column=0, padx=20, pady=(0, 20), sticky="nsew")

    def update_watermark_info(self, path):
        if path and os.path.isfile(path):
            try:
                with Image.open(path) as img: w, h = img.size; self.watermark_dims.set(f"Size: {w}x{h} px")
            except Exception as e: self.watermark_dims.set("Size: Read Error"); print(f"Error reading watermark size '{path}': {e}")
        else: self.watermark_dims.set("Size: (not selected)")

    def select_main_folder(self):
        path = filedialog.askdirectory(title="Select MAIN folder"); path and self.main_folder.set(path)

    def select_watermark_file(self):
        path = filedialog.askopenfilename(title="Select watermark file", filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("PNG", "*.png"), ("JPEG", "*.jpg *.jpeg")])
        if path:
            wm_ext = os.path.splitext(path)[1].lower()
            if wm_ext not in ALLOWED_WATERMARK_EXTENSIONS: tkinter.messagebox.showerror("Format Error", f"Format ({wm_ext}) is not supported."); return
            self.watermark_file.set(path); self.update_watermark_info(path)

    def _browse_magick_path(self):
        if sys.platform == "win32": filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        else: filetypes = [("All files", "*.*")]
        path = filedialog.askopenfilename(title="Specify path to ImageMagick executable ('magick')", filetypes=filetypes)
        if path: self.magick_path_var.set(path)

    def _check_magick_executable(self, path_to_check):
        if not path_to_check or not os.path.isfile(path_to_check): return False, "File not found or path not specified."
        command = [path_to_check, "-version"]
        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True, encoding='utf-8', errors='replace', startupinfo=None, timeout=10)
            print(f"ImageMagick version check successful:\n{result.stdout[:200]}...")
            return True, None
        except FileNotFoundError: return False, f"File '{os.path.basename(path_to_check)}' not found."
        except subprocess.CalledProcessError as e: error_message = f"ImageMagick command returned an error (code {e.returncode})."; print(f"--- IMAGEMAGICK Version Check Error ---\nCommand: {' '.join(e.cmd)}\nStderr:\n{e.stderr}\n--- END ERROR ---"); return False, error_message
        except subprocess.TimeoutExpired: return False, "ImageMagick check timed out."
        except Exception as e: error_message = f"Unexpected error checking ImageMagick: {type(e).__name__}"; print(f"--- UNEXPECTED Check Error ---\n{traceback.format_exc()}\n--- END ERROR ---"); return False, error_message

    def _check_and_save_magick_path(self):
        magick_path = self.magick_path_var.get()
        is_valid, error_msg = self._check_magick_executable(magick_path)
        if is_valid:
            self.verified_magick_path = magick_path
            if hasattr(self, 'psd_radio_button'): self.psd_radio_button.configure(state="normal")
            tkinter.messagebox.showinfo("Success", f"ImageMagick path verified and saved:\n{magick_path}")
            self.save_settings()
        else:
            self.verified_magick_path = None
            if hasattr(self, 'psd_radio_button'): self.psd_radio_button.configure(state="disabled")
            if self.process_type.get() == "psd": self.process_type.set("png")
            tkinter.messagebox.showerror("Error", f"Failed to verify ImageMagick path:\n{error_msg}")

    def _validate_loaded_magick_path(self):
        magick_path = self.magick_path_var.get()
        is_valid, _ = self._check_magick_executable(magick_path)
        if is_valid:
            self.verified_magick_path = magick_path
            if hasattr(self, 'psd_radio_button'): self.psd_radio_button.configure(state="normal")
            print(f"Loaded ImageMagick path verified: {magick_path}")
        else:
            self.verified_magick_path = None
            if hasattr(self, 'psd_radio_button'): self.psd_radio_button.configure(state="disabled")
            if self.process_type.get() == "psd": self.process_type.set("png")
            if magick_path: print(f"Warning: Loaded ImageMagick path is invalid or not working: {magick_path}")

    def update_status(self, message):
        if hasattr(self, 'status_textbox') and self.status_textbox and self.status_textbox.winfo_exists():
            try: self.status_textbox.configure(state="normal"); self.status_textbox.insert("end", str(message) + "\n"); self.status_textbox.see("end"); self.status_textbox.configure(state="disabled"); self.update_idletasks()
            except tkinter.TclError: pass

    def update_progress(self, value):
         if hasattr(self, 'progress_bar') and self.progress_bar and self.progress_bar.winfo_exists():
            try: value = max(0.0, min(float(value), 1.0)); self.progress_bar.set(value); self.update_idletasks()
            except tkinter.TclError: pass

    def enable_controls(self, enable=True):
        new_state = "normal" if enable else "disabled"
        widget_names = ['main_folder_btn', 'watermark_btn', 'freq_entry', 'step_entry', 'thresh_entry', 'max_steps_entry', 'zip_checkbox', 'start_button', 'magick_path_entry', 'magick_browse_btn', 'magick_check_btn', 'png_radio_button', 'psd_radio_button']
        for name in widget_names:
             widget = getattr(self, name, None)
             if widget and widget.winfo_exists():
                 if name == 'psd_radio_button' and enable:
                      widget_state = "normal" if self.verified_magick_path else "disabled"
                      try: widget.configure(state=widget_state)
                      except tkinter.TclError: pass
                 else:
                      try: widget.configure(state=new_state)
                      except tkinter.TclError: pass
        start_button = getattr(self, 'start_button', None)
        if start_button and start_button.winfo_exists():
             try: start_button.configure(text="Start Processing" if enable else "Processing...")
             except tkinter.TclError: pass

    def start_processing_thread(self):
        base_input_dir = self.main_folder.get(); watermark_path = self.watermark_file.get()
        if not base_input_dir or not os.path.isdir(base_input_dir): tkinter.messagebox.showerror("Error", "Please select the main folder."); return
        if not watermark_path or not os.path.isfile(watermark_path): tkinter.messagebox.showerror("Error", "Please select the watermark file."); return
        wm_ext = os.path.splitext(watermark_path)[1].lower();
        if wm_ext not in ALLOWED_WATERMARK_EXTENSIONS: tkinter.messagebox.showerror("Error", f"Watermark format ({wm_ext}) is not supported."); return
        selected_process_type = self.process_type.get()
        if selected_process_type == "psd" and not self.verified_magick_path: tkinter.messagebox.showerror("Error", "Processing PSD/PSB requires specifying and verifying the ImageMagick path."); return
        valid_numbers = True; config_values = {}; checks = {"Frequency": (self.frequency, 1), "Search step": (self.search_step, 1), "Uniformity thresh.": (self.threshold, 0), "Max steps": (self.max_steps, 0)}
        for name, (var, min_val) in checks.items():
            val_str = var.get(); key = name.lower().replace(" ", "_").replace(".","")
            if not is_int(val_str) or int(val_str) < min_val: tkinter.messagebox.showerror("Input Error", f"{name} must be a number {'>=' if min_val >= 0 else '>'} {min_val}."); valid_numbers = False; break
            else: config_values[key] = int(val_str)
        if not valid_numbers: return
        config = {'frequency': config_values["frequency"], 'search_step': config_values["search_step"], 'threshold': config_values["uniformity_thresh"], 'max_steps': config_values["max_steps"], 'create_zip': self.create_zip.get()}
        self.enable_controls(False); self.progress_bar.set(0)
        log_textbox = getattr(self, 'status_textbox', None);
        if log_textbox and log_textbox.winfo_exists(): log_textbox.configure(state="normal"); log_textbox.delete("1.0", "end"); log_textbox.configure(state="disabled")
        self.update_status(f"Folder: {base_input_dir}"); self.update_status(f"Watermark: {os.path.basename(watermark_path)}"); self.update_status(f"File Type: {selected_process_type.upper()}"); self.update_status(f"ZIP Mode: {'On' if config['create_zip'] else 'Off'}"); self.update_status("--- Start ---")
        magick_exe_to_use = self.verified_magick_path if selected_process_type == "psd" else DEFAULT_IMAGEMAGICK_COMMAND
        processing_thread = threading.Thread(target=self.run_processing, args=(base_input_dir, watermark_path, selected_process_type, magick_exe_to_use, config), daemon=True); processing_thread.start()

    def run_processing(self, base_input_dir, watermark_path, selected_process_type, magick_executable, config):
        main_output_dir = base_input_dir.rstrip('/\\') + OUTPUT_SUFFIX
        try: os.makedirs(main_output_dir, exist_ok=True)
        except OSError as e: self.after(0, self.update_status, f"! Critical error creating folder '{main_output_dir}': {e}"); self.after(0, self.enable_controls, True); return

        extensions_to_process = EXTENSIONS_PSD_PSB if selected_process_type == "psd" else EXTENSIONS_PNG_JPG
        magick_exe_path = magick_executable

        folders_to_process = []; single_folder_files = []; has_subfolders = False
        try:
            for item_name in os.listdir(base_input_dir):
                item_full_path = os.path.join(base_input_dir, item_name)
                if os.path.isdir(item_full_path) and item_full_path.rstrip('/\\') != main_output_dir: folders_to_process.append(item_full_path); has_subfolders = True
            if not has_subfolders:
                for filename in os.listdir(base_input_dir):
                    file_full_path = os.path.join(base_input_dir, filename)
                    if filename.lower().endswith(extensions_to_process):
                        if os.path.abspath(file_full_path) != os.path.abspath(watermark_path): single_folder_files.append(filename)
                if not single_folder_files: self.after(0, self.update_status, f"! No files of type {selected_process_type.upper()} found in folder."); self.after(0, self.enable_controls, True); return
                folders_to_process = [base_input_dir]; self.after(0, self.update_status, f"Found {len(single_folder_files)} files ({selected_process_type.upper()}) in base folder.")
            else: self.after(0, self.update_status, f"Found {len(folders_to_process)} subfolders to process.")
        except Exception as scan_error: self.after(0, self.update_status, f"! Error reading folder '{base_input_dir}': {scan_error}"); self.after(0, self.enable_controls, True); return

        total_folders = len(folders_to_process); total_files_processed_successfully = 0; total_files_with_errors = 0

        for folder_index, current_folder_path in enumerate(folders_to_process):
            current_folder_name = os.path.basename(current_folder_path)
            self.after(0, self.update_status, f"\n[{folder_index+1}/{total_folders}] Folder: {current_folder_name}")
            output_path = None; is_zip_mode = config['create_zip']; zip_file_object = None
            if is_zip_mode:
                output_path = os.path.join(main_output_dir, current_folder_name + ".zip")
                try: zip_file_object = zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED)
                except Exception as zip_create_error: self.after(0, self.update_status, f" ! ZIP Error '{output_path}': {zip_create_error}"); total_files_with_errors += 1; continue
            else:
                output_path = os.path.join(main_output_dir, current_folder_name)
                try: os.makedirs(output_path, exist_ok=True)
                except OSError as dir_create_error: self.after(0, self.update_status, f" ! Folder Error '{output_path}': {dir_create_error}"); total_files_with_errors += 1; continue

            files_to_process_in_folder = []
            if not has_subfolders: files_to_process_in_folder = single_folder_files
            else:
                try:
                    for filename in os.listdir(current_folder_path):
                         if filename.lower().endswith(extensions_to_process):
                             file_full_path = os.path.join(current_folder_path, filename)
                             if os.path.abspath(file_full_path) != os.path.abspath(watermark_path): files_to_process_in_folder.append(filename)
                except Exception as listdir_error:
                    self.after(0, self.update_status, f" ! Error reading files from '{current_folder_path}': {listdir_error}")
                    if zip_file_object:
                        try:
                            zip_file_object.close()
                            if os.path.exists(output_path):
                                os.remove(output_path)
                                self.after(0, self.update_status, f" - Removed ZIP due to folder read error: {os.path.basename(output_path)}")
                        except Exception as e_close:
                            print(f"Warning: Error closing/removing zip after listdir error '{output_path}': {e_close}")
                    total_files_with_errors += 1; continue

            number_of_files = len(files_to_process_in_folder)
            if number_of_files == 0:
                self.after(0, self.update_status, f" - No files of type {selected_process_type.upper()} found.")
                if zip_file_object:
                     try:
                         zip_file_object.close()
                         if os.path.exists(output_path):
                             os.remove(output_path)
                             self.after(0, self.update_status, f" - Removed empty ZIP: {os.path.basename(output_path)}")
                     except Exception as e_close:
                          print(f"Warning: Error closing/removing empty zip '{output_path}': {e_close}")
                continue

            folder_success_files = 0; folder_error_files = 0; last_processed_file_index = -1
            try:
                with tempfile.TemporaryDirectory(prefix="awm_", dir=main_output_dir) as temp_conversion_dir:
                    for file_index, current_filename in enumerate(files_to_process_in_folder):
                        last_processed_file_index = file_index
                        original_file_path = os.path.join(current_folder_path, current_filename)
                        self.after(0, self.update_status, f" >> {current_filename}")

                        temp_png_path = convert_to_temp_png(original_file_path, temp_conversion_dir, magick_exe_path, lambda msg: self.after(0, self.update_status, msg))
                        current_file_success = False
                        if temp_png_path and os.path.exists(temp_png_path):

                            output_png_filename = os.path.splitext(current_filename)[0] + ".png"
                            final_destination_path_or_arcname = None; path_for_watermarked_output = None
                            if is_zip_mode:
                                 final_destination_path_or_arcname = output_png_filename
                                 path_for_watermarked_output = os.path.join(temp_conversion_dir, "_marked_" + output_png_filename)
                            else:
                                 final_destination_path_or_arcname = os.path.join(output_path, output_png_filename)
                                 path_for_watermarked_output = final_destination_path_or_arcname
                            watermark_step_success = add_watermarks_to_image(temp_png_path, watermark_path, path_for_watermarked_output, config, lambda msg: self.after(0, self.update_status, msg))

                            if watermark_step_success and is_zip_mode and zip_file_object:
                                try:
                                    zip_file_object.write(path_for_watermarked_output, final_destination_path_or_arcname)
                                    os.remove(path_for_watermarked_output)
                                except Exception as zip_write_error:
                                    self.after(0, self.update_status, f"  ! Error adding to ZIP {final_destination_path_or_arcname}: {zip_write_error}"); watermark_step_success = False
                                    if os.path.exists(path_for_watermarked_output): os.remove(path_for_watermarked_output)
                            current_file_success = watermark_step_success
                        else: current_file_success = False

                        if current_file_success: folder_success_files += 1
                        else: folder_error_files += 1

                        progress_in_folder = (file_index + 1) / number_of_files
                        overall_progress = (folder_index + progress_in_folder) / total_folders
                        self.after(0, self.update_progress, overall_progress)
            except Exception as folder_processing_error:
                 self.after(0, self.update_status, f"! Critical error processing folder {current_folder_name}: {folder_processing_error}")
                 remaining_files = number_of_files - last_processed_file_index - 1
                 folder_error_files += remaining_files
                 total_files_with_errors += remaining_files

            log_suffix = f"Success: {folder_success_files}" + (f", Errors: {folder_error_files}" if folder_error_files > 0 else "")
            self.after(0, self.update_status, f"   {log_suffix}")
            total_files_processed_successfully += folder_success_files; total_files_with_errors += folder_error_files
            if zip_file_object:
                try:
                    zip_file_object.close()
                    if folder_error_files == number_of_files and number_of_files > 0 and os.path.exists(output_path):
                         self.after(0, self.update_status, f" - Removed erroneous ZIP: {os.path.basename(output_path)}")
                         os.remove(output_path)
                except Exception as zip_close_error: self.after(0, self.update_status, f" ! Error closing ZIP {os.path.basename(output_path)}: {zip_close_error}")

        self.after(0, self.update_status, f"\n--- Done. Success: {total_files_processed_successfully}, Errors: {total_files_with_errors} ---")
        self.after(0, self.update_progress, 1.0); self.after(0, self.enable_controls, True)

    def load_settings(self):
        config_path = get_config_path(); settings = DEFAULT_CONFIG.copy()
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    for key, default_value in DEFAULT_CONFIG.items(): settings[key] = loaded_settings.get(key, default_value)
                print(f"Settings loaded from {config_path}")
            else: print("Settings file not found.")
        except Exception as e: print(f"Error loading settings from '{config_path}': {e}."); settings = DEFAULT_CONFIG.copy()

        main_folder_path = settings.get("main_folder", ""); self.main_folder.set(main_folder_path if main_folder_path and os.path.isdir(main_folder_path) else "")
        watermark_file_path = settings.get("watermark_file", "");
        if watermark_file_path and os.path.isfile(watermark_file_path): self.watermark_file.set(watermark_file_path); self.update_watermark_info(watermark_file_path)
        else: self.watermark_file.set(""); self.update_watermark_info(None)
        self.frequency.set(str(settings.get("frequency", DEFAULT_CONFIG["frequency"]))); self.search_step.set(str(settings.get("search_step", DEFAULT_CONFIG["search_step"])))
        self.threshold.set(str(settings.get("threshold", DEFAULT_CONFIG["threshold"]))); self.max_steps.set(str(settings.get("max_steps", DEFAULT_CONFIG["max_steps"])))
        self.create_zip.set(bool(settings.get("create_zip", DEFAULT_CONFIG["create_zip"])))
        self.magick_path_var.set(settings.get("magick_path", ""))
        self.process_type.set(settings.get("process_type", "png"))

    def save_settings(self):
        settings = {"main_folder": self.main_folder.get(), "watermark_file": self.watermark_file.get(), "frequency": self.frequency.get(), "search_step": self.search_step.get(), "threshold": self.threshold.get(), "max_steps": self.max_steps.get(), "create_zip": self.create_zip.get(),
                    "magick_path": self.magick_path_var.get(),
                    "process_type": self.process_type.get()
                   }
        config_path = get_config_path();
        try:
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f: json.dump(settings, f, indent=4, ensure_ascii=False)
            print(f"Settings saved to {config_path}")
        except Exception as e: print(f"Error saving settings to '{config_path}': {e}")

    def on_closing(self):
        print("Window closing..."); self.save_settings(); self.destroy()

if __name__ == "__main__":
    if sys.platform == "win32":
        try: from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1); print("DPI Awareness OK.")
        except Exception as e: print(f"DPI awareness failed: {e}")

    app = WatermarkerApp()
    app.mainloop()