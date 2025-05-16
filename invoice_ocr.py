#!/usr/bin/env python3
"""
Invoice OCR Tool - Convert scanned invoices to Excel
This tool extracts tabular data from scanned PDFs and JPEGs and exports to Excel.
"""
import os
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from pathlib import Path
import threading
import re

import cv2
import numpy as np
import pytesseract
import xlsxwriter
from pdf2image import convert_from_path
from PIL import Image, ImageTk, ImageDraw

class ImageRegionSelectionDialog(tk.Toplevel):
    """Dialog for selecting fields by drawing regions on an image."""
    
    MAX_CANVAS_WIDTH = 700
    MAX_CANVAS_HEIGHT = 550

    def __init__(self, parent, pil_image, tesseract_config, initial_fields=None):
        super().__init__(parent)
        self.parent = parent
        self.pil_image_original = pil_image
        self.tesseract_config = tesseract_config
        self.selected_regions = initial_fields if initial_fields else {} # Format: {"name": {"page_index":0, "bbox":[x1,y1,x2,y2]}}

        self.img_scale = 1.0
        self.canvas_img_display = None # PhotoImage for canvas
        self.current_rect_id = None
        self.start_x = None
        self.start_y = None
        self.found_label_rect_ids = [] # To store IDs of highlighted text rectangles

        self.title("Image Region Field Selection")
        self.geometry("1000x700") # Increased size
        self.minsize(800, 600)
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.load_image_on_canvas()
        self.populate_fields_tree()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Instructions
        instr_text = "Draw a rectangle on the image for the field's VALUE, enter a field name, then click 'Add Field'.\nOptionally, type a field name and click 'Find Text' to highlight it on the image first."
        ttk.Label(main_frame, text=instr_text, font=('Arial', 10)).pack(anchor=tk.W, pady=(0,10))

        # Paned window for image canvas and field list/controls
        pw = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        pw.pack(fill=tk.BOTH, expand=True)

        # Left: Image Canvas
        canvas_frame = ttk.LabelFrame(pw, text="Invoice Image (First Page)")
        pw.add(canvas_frame, weight=3)

        self.canvas = tk.Canvas(canvas_frame, bg="lightgray", scrollregion=(0,0,0,0))
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        
        # Right: Controls and Field List
        controls_frame = ttk.Frame(pw) # No LabelFrame needed
        pw.add(controls_frame, weight=1)

        # Field Name Input
        field_input_frame = ttk.Frame(controls_frame)
        field_input_frame.pack(fill=tk.X, pady=5)
        ttk.Label(field_input_frame, text="Field Name:").pack(side=tk.LEFT, padx=(0,5))
        self.field_name_var = tk.StringVar()
        self.field_name_entry = ttk.Entry(field_input_frame, textvariable=self.field_name_var)
        self.field_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.find_text_button = ttk.Button(field_input_frame, text="Find Text", command=self.find_text_on_image)
        self.find_text_button.pack(side=tk.LEFT, padx=(5,0))

        # Add Field Button
        self.add_field_button = ttk.Button(controls_frame, text="Add Current Region as Field", command=self.add_field_region, state=tk.DISABLED)
        self.add_field_button.pack(fill=tk.X, pady=5)

        # Selected Fields List (Treeview)
        fields_list_frame = ttk.LabelFrame(controls_frame, text="Defined Fields")
        fields_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.fields_tree = ttk.Treeview(fields_list_frame, columns=("name", "bbox"), show="headings")
        self.fields_tree.heading("name", text="Field Name")
        self.fields_tree.heading("bbox", text="Region (x1,y1,x2,y2)")
        self.fields_tree.column("name", width=100, stretch=tk.NO)
        self.fields_tree.column("bbox", width=150)
        self.fields_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        fields_scroll = ttk.Scrollbar(fields_list_frame, orient=tk.VERTICAL, command=self.fields_tree.yview)
        fields_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.fields_tree.config(yscrollcommand=fields_scroll.set)

        # Action buttons for Treeview
        tree_buttons_frame = ttk.Frame(controls_frame)
        tree_buttons_frame.pack(fill=tk.X, pady=5)
        ttk.Button(tree_buttons_frame, text="Remove Selected", command=self.remove_selected_field).pack(side=tk.LEFT, padx=2)
        ttk.Button(tree_buttons_frame, text="Test Extract", command=self.test_extract_selected).pack(side=tk.LEFT, padx=2)

        # Dialog Buttons (OK, Cancel)
        dialog_buttons_frame = ttk.Frame(main_frame)
        dialog_buttons_frame.pack(fill=tk.X, pady=(10,0))
        ttk.Button(dialog_buttons_frame, text="OK", command=self.on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(dialog_buttons_frame, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)

    def load_image_on_canvas(self):
        orig_w, orig_h = self.pil_image_original.size
        
        scale_w = self.MAX_CANVAS_WIDTH / orig_w
        scale_h = self.MAX_CANVAS_HEIGHT / orig_h
        self.img_scale = min(scale_w, scale_h, 1.0) # Don't scale up

        disp_w = int(orig_w * self.img_scale)
        disp_h = int(orig_h * self.img_scale)

        img_for_canvas = self.pil_image_original.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
        self.canvas_img_display = ImageTk.PhotoImage(img_for_canvas)
        
        self.canvas.config(width=disp_w, height=disp_h, scrollregion=(0,0,disp_w,disp_h))
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.canvas_img_display)
        self.canvas.image = self.canvas_img_display # Keep reference

    def _on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.current_rect_id:
            self.canvas.delete(self.current_rect_id)
        self.current_rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)
        self.add_field_button.config(state=tk.DISABLED)

    def _on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        if self.current_rect_id:
            self.canvas.coords(self.current_rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def _on_release(self, event):
        if self.start_x is None or self.start_y is None: # No press initiated
            return

        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        # Ensure x1 < x2 and y1 < y2
        self.rect_coords_canvas = (
            min(self.start_x, end_x),
            min(self.start_y, end_y),
            max(self.start_x, end_x),
            max(self.start_y, end_y)
        )
        # Enable Add Field button if a valid rectangle is drawn
        if abs(self.start_x - end_x) > 5 and abs(self.start_y - end_y) > 5: # Min size for a rect
            try:
                # Autopopulate field name with OCR of the selected region
                x1_orig = int(self.rect_coords_canvas[0] / self.img_scale)
                y1_orig = int(self.rect_coords_canvas[1] / self.img_scale)
                x2_orig = int(self.rect_coords_canvas[2] / self.img_scale)
                y2_orig = int(self.rect_coords_canvas[3] / self.img_scale)
                bbox_original = [x1_orig, y1_orig, x2_orig, y2_orig]

                cropped_img = self.pil_image_original.crop(bbox_original)
                extracted_text_full = pytesseract.image_to_string(cropped_img, config=self.tesseract_config).strip()
                
                # Sanitize the extracted text to be a valid field name
                if extracted_text_full:
                    # Replace newlines with a single space, then normalize multiple spaces to one
                    text_for_name = re.sub(r'\s+', ' ', extracted_text_full.replace('\n', ' ')).strip()
                    
                    sanitized_name = re.sub(r'[^a-zA-Z0-9_\s-]', '', text_for_name).strip() # Keep alphanumeric, underscore, space, hyphen
                    sanitized_name = re.sub(r'\s+', '_', sanitized_name) # Replace spaces with underscore
                    sanitized_name = sanitized_name[:30] # Limit length
                    if sanitized_name:
                        self.field_name_var.set(sanitized_name)
            except Exception as e:
                # Log this error or handle it silently, e.g. print to console for debugging
                print(f"Error during auto-OCR for field name: {e}")
                # Optionally, inform the main app's log if a logger is passed or accessible
                # self.parent.log(f"OCR Error for field name: {e}") 

            self.add_field_button.config(state=tk.NORMAL)
            self.field_name_entry.focus_set()
        else: # Rectangle too small, delete it
            if self.current_rect_id:
                self.canvas.delete(self.current_rect_id)
            self.current_rect_id = None
            self.add_field_button.config(state=tk.DISABLED)
        
        # self.start_x, self.start_y will be reset on next press

    def add_field_region(self):
        if not self.current_rect_id or not self.rect_coords_canvas:
            messagebox.showwarning("No Region", "Please draw a rectangle on the image first.", parent=self)
            return
        
        field_name = self.field_name_var.get().strip()
        if not field_name:
            field_name = simpledialog.askstring("Field Name", "Enter a name for this field:", parent=self)
            if not field_name:
                return # User cancelled or entered nothing
            self.field_name_var.set(field_name)

        if field_name in self.selected_regions:
            messagebox.showwarning("Duplicate Name", f"A field named '{field_name}' already exists.", parent=self)
            return

        # Convert canvas coordinates to original image coordinates
        x1_orig = int(self.rect_coords_canvas[0] / self.img_scale)
        y1_orig = int(self.rect_coords_canvas[1] / self.img_scale)
        x2_orig = int(self.rect_coords_canvas[2] / self.img_scale)
        y2_orig = int(self.rect_coords_canvas[3] / self.img_scale)
        
        bbox_original = [x1_orig, y1_orig, x2_orig, y2_orig]

        self.selected_regions[field_name] = {"page_index": 0, "bbox": bbox_original} # Assuming page 0 for now
        self.populate_fields_tree()
        
        # Clear current selection
        self.field_name_var.set("")
        if self.current_rect_id: # Keep the drawn rect on canvas for visual feedback
            # Change color to indicate it's saved
            self.canvas.itemconfig(self.current_rect_id, outline="blue")
        self.current_rect_id = None # Ready for new rect
        self.rect_coords_canvas = None
        self.add_field_button.config(state=tk.DISABLED)

    def populate_fields_tree(self):
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for name, data in self.selected_regions.items():
            bbox_str = f"({data['bbox'][0]},{data['bbox'][1]},{data['bbox'][2]},{data['bbox'][3]})"
            self.fields_tree.insert("", tk.END, values=(name, bbox_str))

    def remove_selected_field(self):
        selected_item = self.fields_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a field from the list to remove.", parent=self)
            return
        
        field_name = self.fields_tree.item(selected_item[0], "values")[0]
        if field_name in self.selected_regions:
            del self.selected_regions[field_name]
            self.populate_fields_tree()
            # Optionally, remove the corresponding rectangle from canvas if we stored its ID
            # For now, clearing all non-blue (committed) rectangles might be simpler if needed
            # Or, redraw all committed rectangles. For now, just update tree.

    def test_extract_selected(self):
        selected_item = self.fields_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a field to test.", parent=self)
            return

        field_name = self.fields_tree.item(selected_item[0], "values")[0]
        if field_name in self.selected_regions:
            bbox = self.selected_regions[field_name]["bbox"]
            try:
                cropped_img = self.pil_image_original.crop(bbox)
                # Ensure pytesseract is available
                text = pytesseract.image_to_string(cropped_img, config=self.tesseract_config)
                messagebox.showinfo("Test Extraction", f"Field: {field_name}\nExtracted: {text.strip()}", parent=self)
            except Exception as e:
                messagebox.showerror("Extraction Error", f"Could not extract text: {e}", parent=self)
        
    def clear_found_label_highlights(self):
        for rect_id in self.found_label_rect_ids:
            self.canvas.delete(rect_id)
        self.found_label_rect_ids.clear()

    def find_text_on_image(self):
        self.clear_found_label_highlights()
        if self.current_rect_id:  # Clear any existing red selection box
            self.canvas.delete(self.current_rect_id)
            self.current_rect_id = None
            self.rect_coords_canvas = None
            self.add_field_button.config(state=tk.DISABLED)

        search_text_original = self.field_name_var.get().strip()
        if not search_text_original:
            messagebox.showinfo("Input Needed", "Please enter text in the 'Field Name' box to find.", parent=self)
            return

        self.update_status_local(f"Searching for '{search_text_original}'...")
        
        try:
            data = pytesseract.image_to_data(self.pil_image_original,
                                             config=self.tesseract_config,
                                             output_type=pytesseract.Output.DICT)
            n_boxes = len(data['level'])
            if n_boxes == 0:
                messagebox.showinfo("OCR Error", "No text could be detected on the image.", parent=self)
                self.update_status_local("")
                return

            search_text_lower_normalized = " ".join(search_text_original.lower().split())
            search_words_list = search_text_lower_normalized.split()
            num_search_words = len(search_words_list)

            found_labels_details = []  # Store details of all occurrences of the label phrase

            for i in range(n_boxes - num_search_words + 1):
                # Try to match the search_words_list starting at data['text'][i]
                
                # Check basic conditions for the first word of a potential phrase
                first_word_ocr_text = data['text'][i].strip().lower()
                first_word_conf = int(data['conf'][i])
                
                # First word of search phrase must be in the first OCR word, and conf must be good
                if not (search_words_list[0] in first_word_ocr_text and first_word_conf >= 30):
                    continue

                current_phrase_ocr_words_info = [] 
                is_valid_phrase_so_far = True
                
                # Validate the sequence of words to form the searched phrase
                for k_word_in_phrase in range(num_search_words):
                    current_ocr_word_idx = i + k_word_in_phrase
                    if current_ocr_word_idx >= n_boxes: # Bounds check
                        is_valid_phrase_so_far = False
                        break

                    ocr_word_text = data['text'][current_ocr_word_idx].strip()
                    ocr_word_text_lower = ocr_word_text.lower()
                    ocr_word_conf = int(data['conf'][current_ocr_word_idx])
                    
                    ocr_word_left = data['left'][current_ocr_word_idx]
                    ocr_word_top = data['top'][current_ocr_word_idx]
                    ocr_word_width = data['width'][current_ocr_word_idx]
                    ocr_word_height = data['height'][current_ocr_word_idx]

                    if ocr_word_conf < 20:  # Min confidence for any word in the phrase
                        is_valid_phrase_so_far = False
                        break
                    
                    # Check if current OCR word matches the corresponding search word (substring match)
                    if search_words_list[k_word_in_phrase] not in ocr_word_text_lower:
                        is_valid_phrase_so_far = False
                        break
                    
                    word_info = {'text': ocr_word_text, 'x': ocr_word_left, 'y': ocr_word_top, 
                                 'w': ocr_word_width, 'h': ocr_word_height, 'idx': current_ocr_word_idx}

                    if k_word_in_phrase > 0: # For subsequent words, check spatial coherence
                        prev_word_info = current_phrase_ocr_words_info[-1]
                        
                        # Vertical alignment: roughly same line (top coordinates should be similar)
                        # Allow deviation up to 75% of the previous word's height
                        if abs(ocr_word_top - prev_word_info['y']) > prev_word_info['h'] * 0.75:
                            is_valid_phrase_so_far = False
                            break
                        
                        # Horizontal alignment: current word must be to the right of previous, and not too far
                        # Max gap can be roughly 1.5 times the height of the previous word (generous for spaces)
                        max_gap = prev_word_info['h'] * 1.5 
                        actual_gap = ocr_word_left - (prev_word_info['x'] + prev_word_info['w'])
                        if not (0 <= actual_gap <= max_gap) : # Must be to the right, within reasonable gap
                            is_valid_phrase_so_far = False
                            break
                    
                    current_phrase_ocr_words_info.append(word_info)

                if is_valid_phrase_so_far and current_phrase_ocr_words_info:
                    # Final check: the assembled phrase should match the search query (normalized)
                    formed_phrase_text = " ".join(p['text'] for p in current_phrase_ocr_words_info)
                    normalized_formed_phrase = " ".join(re.sub(r'[^a-z0-9]', '', p.lower()) for p in formed_phrase_text.split() if re.sub(r'[^a-z0-9]', '', p.lower()))
                    normalized_search_text = " ".join(re.sub(r'[^a-z0-9]', '', sw.lower()) for sw in search_words_list if re.sub(r'[^a-z0-9]', '', sw.lower()))

                    if normalized_search_text in normalized_formed_phrase:
                        label_x1 = min(w['x'] for w in current_phrase_ocr_words_info)
                        label_y1 = min(w['y'] for w in current_phrase_ocr_words_info)
                        label_x2 = max(w['x'] + w['w'] for w in current_phrase_ocr_words_info)
                        label_y2 = max(w['y'] + w['h'] for w in current_phrase_ocr_words_info)
                        
                        found_labels_details.append({
                            'text': formed_phrase_text,
                            'bbox_orig': [label_x1, label_y1, label_x2, label_y2],
                            'indices': [w['idx'] for w in current_phrase_ocr_words_info]
                        })
            
            label_found_count = len(found_labels_details)
            value_auto_selected = False
            auto_selected_value_text = ""

            if label_found_count > 0:
                for label_info in found_labels_details: # Highlight all found label instances
                    lx1_orig, ly1_orig, lx2_orig, ly2_orig = label_info['bbox_orig']
                    canvas_lx1, canvas_ly1 = lx1_orig * self.img_scale, ly1_orig * self.img_scale
                    canvas_lx2, canvas_ly2 = lx2_orig * self.img_scale, ly2_orig * self.img_scale
                    rect_id = self.canvas.create_rectangle(
                        canvas_lx1, canvas_ly1, canvas_lx2, canvas_ly2,
                        outline="yellow", width=2, dash=(4, 4)
                    )
                    self.found_label_rect_ids.append(rect_id)

                # Attempt to find value for the first valid label found
                # (Could be extended to let user pick if multiple labels are found)
                for label_info in found_labels_details:
                    if value_auto_selected: break 

                    label_bbox_orig = label_info['bbox_orig']
                    label_indices = label_info['indices']
                    lx1_orig, ly1_orig, lx2_orig, ly2_orig = label_bbox_orig
                    
                    label_width_orig = lx2_orig - lx1_orig
                    label_height_orig = ly2_orig - ly1_orig
                    label_y_mid_orig = ly1_orig + label_height_orig / 2.0
                    
                    search_x_start_after_label_orig = lx2_orig
                    search_x_end_limit_orig = search_x_start_after_label_orig + min(label_width_orig * 4, 400) # Increased search width for value

                    candidate_value_words = []
                    for j in range(n_boxes):
                        if j in label_indices: continue # Value word cannot be part of the label itself

                        word_x_orig, word_y_orig = data['left'][j], data['top'][j]
                        word_w_orig, word_h_orig = data['width'][j], data['height'][j]
                        word_text_j = data['text'][j].strip()
                        word_conf_j = int(data['conf'][j])

                        if not word_text_j or word_conf_j < 15: continue # Lower conf for value words is okay

                        word_y_mid_orig = word_y_orig + word_h_orig / 2.0
                        
                        is_to_right = (word_x_orig >= search_x_start_after_label_orig - 10) and \
                                      (word_x_orig < search_x_end_limit_orig) # Allow small overlap/start slightly left
                        
                        vertical_tolerance_factor = 0.85 # Increased from 0.75
                        is_vertically_aligned = abs(word_y_mid_orig - label_y_mid_orig) < \
                                                (label_height_orig + word_h_orig) / 2.0 * vertical_tolerance_factor

                        if is_to_right and is_vertically_aligned:
                            candidate_value_words.append({
                                'text': word_text_j, 'x': word_x_orig, 'y': word_y_orig, 
                                'w': word_w_orig, 'h': word_h_orig
                            })
                    
                    if candidate_value_words:
                        candidate_value_words.sort(key=lambda w: (w['y'], w['x'])) # Sort by position
                        
                        # Further filter: group words that are close together to form the value phrase
                        final_value_words = []
                        if candidate_value_words:
                            final_value_words.append(candidate_value_words[0])
                            for k_val in range(1, len(candidate_value_words)):
                                prev_val_word = final_value_words[-1]
                                curr_val_word = candidate_value_words[k_val]
                                # Check if current value word is close to the previous one
                                gap_val = curr_val_word['x'] - (prev_val_word['x'] + prev_val_word['w'])
                                vertical_diff_val = abs(curr_val_word['y'] - prev_val_word['y'])
                                if gap_val < prev_val_word['h'] * 1.0 and vertical_diff_val < prev_val_word['h'] * 0.5 : # Max gap = 1x height, small vertical diff
                                    final_value_words.append(curr_val_word)
                                else:
                                    # If a significant break, consider if this word starts a new potential value line
                                    # For now, we take the first contiguous block
                                    break 
                        
                        if final_value_words:
                            val_x1_orig = min(w['x'] for w in final_value_words)
                            val_y1_orig = min(w['y'] for w in final_value_words)
                            val_x2_orig = max(w['x'] + w['w'] for w in final_value_words)
                            val_y2_orig = max(w['y'] + w['h'] for w in final_value_words)
                            
                            auto_selected_value_text = " ".join(w['text'] for w in final_value_words)

                            if auto_selected_value_text:
                                cvx1, cvy1 = val_x1_orig * self.img_scale, val_y1_orig * self.img_scale
                                cvx2, cvy2 = val_x2_orig * self.img_scale, val_y2_orig * self.img_scale

                                if self.current_rect_id: self.canvas.delete(self.current_rect_id)
                                
                                self.start_x, self.start_y = cvx1, cvy1 
                                self.rect_coords_canvas = (min(cvx1,cvx2), min(cvy1,cvy2), max(cvx1,cvx2), max(cvy1,cvy2))
                                
                                self.current_rect_id = self.canvas.create_rectangle(
                                    self.rect_coords_canvas[0], self.rect_coords_canvas[1],
                                    self.rect_coords_canvas[2], self.rect_coords_canvas[3],
                                    outline="red", width=2
                                )
                                self.add_field_button.config(state=tk.NORMAL)
                                self.canvas.focus_set()
                                value_auto_selected = True
            
            # Final message based on what was found
            if value_auto_selected:
                messagebox.showinfo("Value Auto-Selected", 
                                    f"Found label '{search_text_original}' and a potential value region (text: '{auto_selected_value_text[:70]}...').\n" 
                                    f"The red box has been placed. Adjust if needed, or click 'Add Current Region as Field'.", 
                                    parent=self)
            elif label_found_count > 0:
                messagebox.showinfo("Text Found", f"Found '{search_text_original}' at {label_found_count} location(s). Highlighted in yellow.\n" 
                                                  "Could not auto-detect a value. Please draw a box around the field's VALUE.", parent=self)
            else:
                messagebox.showinfo("Text Not Found", f"Could not find '{search_text_original}' on the image. " 
                                                      "You can still manually enter the field name and draw its region.", parent=self)

        except Exception as e:
            messagebox.showerror("OCR Error", f"Error during text search: {e}", parent=self)
            print(f"Error in find_text_on_image: {e}\n{traceback.format_exc()}")
        finally:
            self.update_status_local("")

    def update_status_local(self, message): # Helper for a potential status bar in dialog
        # If you add a status label to this dialog, update it here.
        # For now, this method can be a placeholder.
        # Example: if self.dialog_status_label: self.dialog_status_label.config(text=message)
        if message: # Print to console if no dedicated label
            print(f"Dialog Status: {message}")

    def on_ok(self):
        self.clear_found_label_highlights() # Clear any yellow highlights
        self.destroy()

    def on_cancel(self):
        self.selected_regions = {} # Discard changes
        self.clear_found_label_highlights() # Clear any yellow highlights
        self.destroy()

    def get_selected_regions(self):
        return self.selected_regions

class InvoiceOCR:
    """Main class for the Invoice OCR application"""
    
    def __init__(self):
        self.root = None
        self.input_file = ""
        self.output_dir = ""
        self.progress_var = None
        self.status_label = None
        self.preview_label = None
        self.preview_panel = None
        self.log_text = None
        self.extracted_text = "" # Still useful for full text dump
        self.selected_fields = {} # Will store region data: {"name": {"page_index":0, "bbox":[x1,y1,x2,y2]}
        
        # Default settings
        self.settings = {
            'dpi': 300,
            'min_line_length': 100,
            'line_threshold': 50,
            'block_padding': 5,
            'tesseract_config': '--oem 3 --psm 6'
        }
        self.process_button = None
        self.empty_image = None # Will be initialized in create_gui or after root exists

    def log(self, message):
        """Add message to the log text area"""
        if self.log_text:
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)

    def update_status(self, message, progress_value=None):
        """Update status label and progress bar"""
        if self.status_label:
            self.status_label.config(text=message)
        if self.progress_var is not None and progress_value is not None:
            self.progress_var.set(progress_value)
        if self.root:
            self.root.update_idletasks()

    def browse_input(self):
        """Open file dialog to select input file"""
        file_path = filedialog.askopenfilename(
            title="Select Invoice File",
            filetypes=(
                ("PDF files", "*.pdf"), 
                ("JPEG files", ("*.jpg", "*.jpeg")), 
                ("PNG files", "*.png"), 
                ("All files", "*.*")
            )
        )
        if file_path:
            self.input_file = file_path
            self.input_file_var.set(file_path)
            self.log(f"Input file selected: {file_path}")
            self.load_preview(file_path)

    def browse_output(self):
        """Open file dialog to select output directory"""
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.output_dir = dir_path
            self.output_dir_var.set(dir_path)
            self.log(f"Output directory selected: {dir_path}")

    def load_preview(self, file_path):
        """Load a preview of the selected file (first page for PDF)"""
        try:
            self.update_status("Loading preview...", 0)
            if not self.empty_image: # Ensure empty_image is initialized
                self.empty_image = tk.PhotoImage(width=1, height=1)

            if not file_path:
                self.preview_label.config(text="No file selected", image=self.empty_image)
                self.preview_label.image = self.empty_image  # Keep a reference
                return

            file_ext = os.path.splitext(file_path)[1].lower()
            img_preview = None

            if file_ext == '.pdf':
                images = convert_from_path(file_path, first_page=1, last_page=1, dpi=72)
                if images:
                    img_preview = images[0]
            elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                img_preview = Image.open(file_path)
            else:
                self.preview_label.config(text="Unsupported file type for preview", image=self.empty_image)
                self.preview_label.image = self.empty_image
                self.log(f"Unsupported file type for preview: {file_ext}")
                return

            if img_preview:
                panel_width = self.preview_panel.winfo_width()
                panel_height = self.preview_panel.winfo_height()

                if panel_width < 2 or panel_height < 2: # Panel not yet sized
                    panel_width = 400 # Default
                    panel_height = 500

                img_preview.thumbnail((panel_width - 10, panel_height - 10), Image.Resampling.LANCZOS)
                
                photo = ImageTk.PhotoImage(img_preview)
                self.preview_label.config(image=photo, text="")
                self.preview_label.image = photo 
                self.log("Preview loaded.")
            else:
                self.preview_label.config(text="Could not load preview", image=self.empty_image)
                self.preview_label.image = self.empty_image
                self.log("Failed to load preview image.")
            self.update_status("Preview loaded", 100)

        except Exception as e:
            self.log(f"Error loading preview: {str(e)}")
            if not self.empty_image: # Ensure empty_image is initialized
                self.empty_image = tk.PhotoImage(width=1, height=1)
            self.preview_label.config(text=f"Error loading preview: {os.path.basename(file_path)}", image=self.empty_image)
            self.preview_label.image = self.empty_image
            self.update_status(f"Error loading preview: {str(e)}", 0)

    def create_gui(self):
        """Create the GUI interface"""
        self.root = tk.Tk()
        self.root.title("Invoice OCR Tool")
        self.root.geometry("900x700")

        # Initialize empty_image here, after root Tk() is created
        self.empty_image = tk.PhotoImage(width=1, height=1)
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="Input/Output", padding="10")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(file_frame, text="Input File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(file_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Settings
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="DPI:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.dpi_var = tk.IntVar(value=self.settings['dpi'])
        ttk.Spinbox(settings_frame, from_=100, to=600, increment=100, textvariable=self.dpi_var, width=5).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="Line Threshold:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(20, 0))
        self.line_threshold_var = tk.IntVar(value=self.settings['line_threshold'])
        ttk.Spinbox(settings_frame, from_=10, to=200, increment=10, textvariable=self.line_threshold_var, width=5).grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Field selection button
        ttk.Button(settings_frame, text="Select Fields for Excel", command=self.open_field_selection).grid(row=0, column=4, padx=(20, 5), pady=5)
        
        # Preview and log area
        preview_log_frame = ttk.Frame(main_frame)
        preview_log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left side - Preview
        preview_frame = ttk.LabelFrame(preview_log_frame, text="Preview")
        preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.preview_panel = ttk.Frame(preview_frame)
        self.preview_panel.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Initialize preview_label with text only first
        self.preview_label = ttk.Label(self.preview_panel, text="No file selected")
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        # Now configure the image and compound, and keep reference
        self.preview_label.config(image=self.empty_image, compound=tk.LEFT)
        self.preview_label.image = self.empty_image # Keep a reference
        
        # Right side - Log
        log_frame = ttk.LabelFrame(preview_log_frame, text="Log")
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, width=40, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        log_scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        
        # Progress and status
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(bottom_frame, variable=self.progress_var, length=100, mode="determinate")
        progress_bar.pack(fill=tk.X, side=tk.TOP, padx=5, pady=5)
        
        self.status_label = ttk.Label(bottom_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.process_button = ttk.Button(button_frame, text="Process", command=self.start_processing)
        self.process_button.pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)
    
    def open_field_selection(self):
        """Open dialog to select fields by drawing regions on the first page image."""
        if not self.input_file:
            messagebox.showwarning("Input Required", "Please select an input file first.", parent=self.root)
            return

        try:
            self.update_status("Loading image for field selection...", 0)
            pil_image_page = None
            file_ext = os.path.splitext(self.input_file)[1].lower()

            if file_ext == '.pdf':
                images = convert_from_path(self.input_file, dpi=self.settings.get('dpi', 300), first_page=1, last_page=1)
                if images:
                    pil_image_page = images[0]
            elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                pil_image_page = Image.open(self.input_file)
            
            if not pil_image_page:
                messagebox.showerror("Image Load Error", "Could not load the first page of the document for field selection.", parent=self.root)
                self.update_status("Error loading image.", 0)
                return
            
            self.update_status("Image loaded. Opening field selection dialog...", 50)
            self.show_field_selection_dialog(pil_image_page)
            self.update_status("Field selection dialog closed.", 100)

        except Exception as e:
            self.log(f"ERROR opening field selection: {str(e)}")
            messagebox.showerror("Error", f"Could not open field selection: {str(e)}", parent=self.root)
            self.update_status(f"Error: {str(e)}", 0)

    def extract_text_for_field_selection(self):
        """This method is no longer directly used by open_field_selection's primary path."""
        self.log("extract_text_for_field_selection is likely obsolete with image region selection.")
        pass

    def show_field_selection_dialog(self, pil_image_page):
        """Show the image region field selection dialog."""
        dialog = ImageRegionSelectionDialog(self.root, pil_image_page, 
                                            tesseract_config=self.settings.get('tesseract_config', '--oem 3 --psm 6'),
                                            initial_fields=self.selected_fields)
        self.root.wait_window(dialog)
        
        self.selected_fields = dialog.get_selected_regions()
        
        if self.selected_fields:
            field_count = len(self.selected_fields)
            self.log(f"Selected {field_count} fields/regions for Excel export.")
            self.update_status(f"{field_count} fields/regions selected.", None)
        else:
            self.log("No fields/regions selected or selection cancelled.")
            self.update_status("No fields/regions selected.", None)

    def start_processing(self):
        """Start the OCR processing in a new thread"""
        if not self.input_file:
            messagebox.showwarning("Input Required", "Please select an input file first.")
            return
        if not self.output_dir:
            messagebox.showwarning("Output Required", "Please select an output directory.")
            return

        if self.process_button:
            self.process_button.config(state=tk.DISABLED)

        self.log("Starting OCR process...")
        self.update_status("Processing...", 0)

        processing_thread = threading.Thread(target=self.process_invoice)
        processing_thread.daemon = True
        processing_thread.start()

    def process_invoice(self):
        """Main processing function. Extracts text based on selected regions if available."""
        try:
            self.log("Starting invoice processing...")
            if not self.input_file or not self.output_dir:
                messagebox.showerror("Missing Info", "Input file or output directory not set.", parent=self.root)
                return

            file_ext = os.path.splitext(self.input_file)[1].lower()
            base_name = os.path.basename(self.input_file)
            file_name_without_ext = os.path.splitext(base_name)[0]
            
            excel_output = os.path.join(self.output_dir, f"{file_name_without_ext}_fields.xlsx")
            txt_output = os.path.join(self.output_dir, f"{file_name_without_ext}_fulltext.txt")
            
            images_pil = []
            if file_ext == '.pdf':
                self.update_status("Converting PDF to images...", 10)
                images_pil = self.pdf_to_images(self.input_file, dpi=self.settings.get('dpi',300))
            else:
                self.update_status("Loading image...", 10)
                img = Image.open(self.input_file)
                if img:
                    images_pil = [img]
            
            if not images_pil:
                raise Exception("Failed to load images from input file")

            field_ocr_results = {}
            full_extracted_text_parts = []

            if self.selected_fields:
                self.log("Processing defined field regions...")
                self.update_status("Extracting text from defined regions...", 30)
                for page_idx, pil_img in enumerate(images_pil):
                    for field_name, region_data in self.selected_fields.items():
                        if region_data.get("page_index") == page_idx:
                            bbox = region_data["bbox"]
                            try:
                                cropped_pil_img = pil_img.crop(bbox)
                                text = pytesseract.image_to_string(cropped_pil_img, config=self.settings['tesseract_config'])
                                field_ocr_results[field_name] = text.strip()
                                self.log(f"Extracted for '{field_name}' (page {page_idx}): '{text.strip()[:50]}...'")
                            except Exception as e:
                                self.log(f"ERROR extracting region for field '{field_name}' on page {page_idx}: {e}")
                                field_ocr_results[field_name] = "[OCR ERROR]"
                
                self.update_status("Region extraction complete. Saving to Excel...", 70)
                self.export_fields_to_excel(field_ocr_results, excel_output)
                self.log(f"Field data exported to: {excel_output}")

            self.update_status("Extracting full text from document...", 80)
            for i, pil_img in enumerate(images_pil):
                self.log(f"Performing full OCR on page {i+1}...")
                page_text = pytesseract.image_to_string(pil_img, config=self.settings['tesseract_config'])
                full_extracted_text_parts.append(page_text)
            
            self.extracted_text = "\n--- Page Break ---\n".join(full_extracted_text_parts)
            with open(txt_output, 'w', encoding='utf-8') as f:
                f.write(self.extracted_text)
            self.log(f"Full extracted text saved to: {txt_output}")
            
            self.update_status("Processing complete!", 100)
            msg_parts = []
            if self.selected_fields and field_ocr_results:
                msg_parts.append(f"Field data saved to {os.path.basename(excel_output)}")
            msg_parts.append(f"Full text saved to {os.path.basename(txt_output)}")
            messagebox.showinfo("Success", "\n".join(msg_parts), parent=self.root)
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            self.log(f"ERROR in process_invoice: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Processing Error", f"An error occurred: {str(e)}", parent=self.root)
        finally:
            if self.process_button:
                self.process_button.config(state=tk.NORMAL)
    
    def pdf_to_images(self, pdf_path, dpi=None):
        """Convert PDF to a list of PIL Images"""
        if dpi is None:
            dpi = self.settings.get('dpi', 300)
        self.log(f"Converting PDF {os.path.basename(pdf_path)} to images at {dpi} DPI...")
        try:
            images = convert_from_path(pdf_path, dpi=dpi)
            self.log(f"Converted {len(images)} pages from PDF.")
            return images
        except Exception as e:
            self.log(f"Error converting PDF to images: {str(e)}")
            raise

    def preprocess_image(self, image_cv):
        """Preprocess image for OCR (e.g., grayscale, thresholding)"""
        self.log("Preprocessing image...")
        if isinstance(image_cv, Image.Image):
            image_cv = cv2.cvtColor(np.array(image_cv), cv2.COLOR_RGB2BGR)

        gray = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        processed_image = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV, 11, 2
        )
        self.log("Image preprocessing complete.")
        return processed_image

    def detect_tables(self, processed_image):
        """Detect tables in the preprocessed image."""
        self.log("Detecting tables (placeholder)...")
        self.log("Table detection complete (no tables identified by placeholder).")
        return []

    def extract_text(self, image_cv, tables_rois=None):
        """Extract text from the image using Tesseract."""
        self.log("Extracting text from image...")
        if isinstance(image_cv, Image.Image):
            image_cv = cv2.cvtColor(np.array(image_cv), cv2.COLOR_RGB2BGR)

        gray_image = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        
        config = self.settings.get('tesseract_config', '--oem 3 --psm 4')
        
        try:
            text = pytesseract.image_to_string(gray_image, config=config)
            self.log("Text extraction successful.")
            return text
        except Exception as e:
            self.log(f"Error during Tesseract OCR: {str(e)}")
            return ""

    def export_to_excel(self, tables_data, output_file):
        """Export extracted tables data to an Excel file."""
        self.log(f"Exporting data to Excel: {output_file}")
        workbook = xlsxwriter.Workbook(output_file)
        
        if not tables_data:
            worksheet = workbook.add_worksheet("Extracted Text")
            if self.extracted_text:
                worksheet.write(0, 0, "Full Extracted Text:")
                worksheet.write_string(1, 0, self.extracted_text)
                worksheet.set_column(0, 0, 100)
            else:
                worksheet.write(0, 0, "No text was extracted or no tables found.")
            self.log("No table data to export. Saved full extracted text instead (if available).")
        else:
            for i, table in enumerate(tables_data):
                worksheet = workbook.add_worksheet(f"Table_{i+1}")
                for r_idx, row in enumerate(table):
                    for c_idx, cell_text in enumerate(row):
                        worksheet.write(r_idx, c_idx, cell_text)
                self.log(f"Exported Table {i+1} to worksheet.")
        
        try:
            workbook.close()
            self.log("Excel export complete.")
        except Exception as e:
            self.log(f"Error closing Excel workbook: {str(e)}")
            raise

    def export_fields_to_excel(self, field_ocr_results, output_file):
        """Export OCRed field data (from regions) to an Excel file.
        Field names will be headers in the first row, and their values in the second row."""
        self.log(f"Exporting field data to Excel: {output_file}")
        workbook = xlsxwriter.Workbook(output_file)
        worksheet = workbook.add_worksheet('Extracted Fields from Regions')
        
        if field_ocr_results:
            col = 0
            for field_name, value in field_ocr_results.items():
                worksheet.write(0, col, field_name)  # Write field name as header in row 0
                worksheet.write(1, col, value)      # Write corresponding value in row 1
                # Auto-adjust column width based on header and value length
                header_len = len(field_name)
                value_len = len(str(value)) if value is not None else 0
                worksheet.set_column(col, col, max(header_len, value_len) + 2) # Add a little padding
                col += 1
        else:
            worksheet.write(0, 0, "No field regions were defined or processed.")

        try:
            workbook.close()
            self.log("Field data Excel export complete.")
        except Exception as e:
            self.log(f"Error closing field data Excel workbook: {str(e)}")
            raise

    def run(self):
        """Run the application"""
        self.create_gui()
        self.root.mainloop()


if __name__ == "__main__":
    app = InvoiceOCR()

    app.run()