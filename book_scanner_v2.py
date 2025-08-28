import cv2
import numpy as np
from pyzbar import pyzbar
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import threading
import queue
import os
import platform
from PIL import Image, ImageTk
import requests
import json
import re
import time

class BarcodeScanner:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("IPEVO 4K Barcode Scanner - Manual Scan Mode")
        self.root.geometry("1000x700")
        
        # macOS specific styling
        if platform.system() == "Darwin":
            try:
                self.root.tk.call('tk', 'scaling', 1.0)
            except:
                pass
        
        # Initialize variables
        self.camera = None
        self.is_preview_running = False
        self.scanned_codes = []
        self.last_scan_time = 0
        
        # Create GUI
        self.setup_gui()
        
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Camera controls
        camera_frame = ttk.LabelFrame(main_frame, text="Camera Controls", padding="5")
        camera_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.start_btn = ttk.Button(camera_frame, text="Start Camera", command=self.start_camera)
        self.start_btn.grid(row=0, column=0, padx=(0, 5))
        
        self.stop_btn = ttk.Button(camera_frame, text="Stop Camera", command=self.stop_camera, state="disabled")
        self.stop_btn.grid(row=0, column=1, padx=(0, 5))
        
        self.scan_btn = ttk.Button(camera_frame, text="📖 Scan Now", command=self.scan_single_frame, state="disabled")
        self.scan_btn.grid(row=0, column=2, padx=(5, 5))
        
        self.status_label = ttk.Label(camera_frame, text="Camera: Stopped")
        self.status_label.grid(row=0, column=3, padx=(20, 0))
        
        # Instructions
        instructions = ttk.Label(camera_frame, text="• Start Camera → Position book → Click 'Scan Now'", 
                                font=('Arial', 10, 'italic'))
        instructions.grid(row=1, column=0, columnspan=4, pady=(5, 0), sticky="w")
        
        # Camera preview
        preview_frame = ttk.LabelFrame(main_frame, text="Camera Preview", padding="5")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5), pady=(0, 10))
        
        self.preview_label = ttk.Label(preview_frame, text="Camera preview will appear here\n\nPosition book barcode in view and click 'Scan Now'", 
                                     anchor="center", justify="center")
        self.preview_label.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Scanned codes display
        codes_frame = ttk.LabelFrame(main_frame, text="Scanned Books", padding="5")
        codes_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Treeview for displaying codes
        columns = ('Timestamp', 'ISBN', 'Title', 'Author', 'Publisher')
        self.tree = ttk.Treeview(codes_frame, columns=columns, show='headings', height=15)
        
        column_widths = {
            'Timestamp': 120,
            'ISBN': 130, 
            'Title': 250,
            'Author': 180,
            'Publisher': 150
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100))
        
        scrollbar = ttk.Scrollbar(codes_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Export controls
        export_frame = ttk.LabelFrame(main_frame, text="Export Options", padding="5")
        export_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(export_frame, text="📄 Export to CSV", command=self.export_csv).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(export_frame, text="📊 Export to Excel", command=self.export_excel).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(export_frame, text="🗑️ Clear List", command=self.clear_list).grid(row=0, column=2, padx=(0, 5))
        
        # Statistics
        stats_frame = ttk.LabelFrame(main_frame, text="Statistics", padding="5")
        stats_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.stats_label = ttk.Label(stats_frame, text="Total Scanned: 0 | Books Found: 0")
        self.stats_label.grid(row=0, column=0)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)  # Make book list wider
        main_frame.rowconfigure(1, weight=1)
        codes_frame.columnconfigure(0, weight=1)
        codes_frame.rowconfigure(0, weight=1)
        
    def find_ipevo_camera(self):
        """Find IPEVO camera among available cameras"""
        backends = [cv2.CAP_AVFOUNDATION, cv2.CAP_ANY]
        
        for backend in backends:
            for i in range(10):
                try:
                    cap = cv2.VideoCapture(i, backend)
                    if cap.isOpened():
                        ret, frame = cap.read()
                        if ret:
                            # Set reasonable resolution
                            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
                            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
                            return cap, i
                        cap.release()
                except Exception as e:
                    print(f"Error testing camera {i}: {e}")
                    continue
        
        return None, -1
    
    def is_isbn(self, barcode_data):
        """Check if barcode is an ISBN"""
        clean_isbn = re.sub(r'[-\s]', '', barcode_data)
        
        if len(clean_isbn) == 10:
            return clean_isbn.replace('X', '').isdigit() or (clean_isbn[-1] == 'X' and clean_isbn[:-1].isdigit())
        elif len(clean_isbn) == 13:
            return clean_isbn.isdigit() and clean_isbn.startswith(('978', '979'))
        
        return False
    
    def lookup_book_info(self, isbn):
        """Look up book information using APIs"""
        try:
            clean_isbn = re.sub(r'[-\s]', '', isbn)
            
            # Try Open Library API first
            url = f"https://openlibrary.org/api/books?bibkeys=ISBN:{clean_isbn}&jscmd=data&format=json"
            response = requests.get(url, timeout=5)
            
            if response.status_code == 200:
                data = response.json()
                book_key = f"ISBN:{clean_isbn}"
                
                if book_key in data:
                    book_info = data[book_key]
                    
                    title = book_info.get('title', 'Unknown Title')
                    authors = book_info.get('authors', [])
                    author_names = [author.get('name', 'Unknown Author') for author in authors]
                    author_str = ', '.join(author_names) if author_names else 'Unknown Author'
                    publishers = book_info.get('publishers', [])
                    publisher_str = publishers[0].get('name', 'Unknown Publisher') if publishers else 'Unknown Publisher'
                    
                    return {'title': title, 'author': author_str, 'publisher': publisher_str}
            
            # Try Google Books as backup
            google_url = f"https://www.googleapis.com/books/v1/volumes?q=isbn:{clean_isbn}"
            google_response = requests.get(google_url, timeout=5)
            
            if google_response.status_code == 200:
                google_data = google_response.json()
                
                if google_data.get('totalItems', 0) > 0:
                    item = google_data['items'][0]
                    volume_info = item.get('volumeInfo', {})
                    
                    title = volume_info.get('title', 'Unknown Title')
                    authors = volume_info.get('authors', ['Unknown Author'])
                    author_str = ', '.join(authors)
                    publisher = volume_info.get('publisher', 'Unknown Publisher')
                    
                    return {'title': title, 'author': author_str, 'publisher': publisher}
            
            return None
            
        except Exception as e:
            print(f"Error looking up book info: {e}")
            return None
    
    def start_camera(self):
        """Start camera preview mode"""
        if self.is_preview_running:
            return
            
        self.camera, camera_index = self.find_ipevo_camera()
        if self.camera is None:
            messagebox.showerror("Error", "No camera found. Please check camera connection.")
            return
        
        self.is_preview_running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.scan_btn.config(state="normal")
        self.status_label.config(text=f"Camera: Ready to Scan (Camera {camera_index})")
        
        # Start preview thread
        self.preview_thread = threading.Thread(target=self.preview_loop, daemon=True)
        self.preview_thread.start()
    
    def stop_camera(self):
        """Stop camera"""
        if not self.is_preview_running:
            return
            
        self.is_preview_running = False
        if self.camera:
            self.camera.release()
            self.camera = None
        
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.scan_btn.config(state="disabled")
        self.status_label.config(text="Camera: Stopped")
        
        self.preview_label.configure(image="", text="Camera preview will appear here\n\nPosition book barcode in view and click 'Scan Now'")
    
    def preview_loop(self):
        """Camera preview loop"""
        while self.is_preview_running and self.camera:
            ret, frame = self.camera.read()
            if not ret:
                break
            
            self.root.after(0, self.update_preview, frame.copy())
            threading.Event().wait(0.033)  # ~30 FPS
    
    def scan_single_frame(self):
        """Scan current camera frame for barcodes"""
        if not self.camera or not self.is_preview_running:
            return
        
        # Prevent rapid scanning
        current_time = time.time()
        if current_time - self.last_scan_time < 1.0:
            return
        
        self.last_scan_time = current_time
        self.scan_btn.config(state="disabled", text="Scanning...")
        self.status_label.config(text="📸 Scanning for barcodes...")
        
        # Capture frame
        ret, frame = self.camera.read()
        if not ret:
            self.scan_btn.config(state="normal", text="📖 Scan Now")
            return
        
        # Process in background thread
        scan_thread = threading.Thread(target=self.process_scan_frame, args=(frame,), daemon=True)
        scan_thread.start()
    
    def process_scan_frame(self, frame):
        """Process frame for barcode detection"""
        try:
            barcodes = pyzbar.decode(frame)
            
            if not barcodes:
                self.root.after(0, self.scan_complete, "No barcodes found", "warning")
                return
            
            new_books_found = 0
            duplicates_found = 0
            
            for barcode in barcodes:
                barcode_data = barcode.data.decode('utf-8')
                barcode_type = barcode.type
                
                # Check for duplicates
                duplicate = any(existing['barcode'] == barcode_data for existing in self.scanned_codes)
                
                if duplicate:
                    duplicates_found += 1
                    continue
                
                # Process new barcode
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                title = "Unknown Title"
                author = "Unknown Author"
                publisher = "Unknown Publisher"
                
                if self.is_isbn(barcode_data):
                    print(f"Looking up ISBN: {barcode_data}")
                    book_info = self.lookup_book_info(barcode_data)
                    if book_info:
                        title = book_info['title']
                        author = book_info['author']
                        publisher = book_info['publisher']
                        print(f"Found: {title} by {author}")
                
                code_info = {
                    'timestamp': timestamp,
                    'barcode': barcode_data,
                    'type': barcode_type,
                    'title': title,
                    'author': author,
                    'publisher': publisher
                }
                
                self.scanned_codes.append(code_info)
                self.root.after(0, self.add_code_to_tree, code_info)
                self.root.after(0, self.update_stats)
                new_books_found += 1
            
            # Show results
            if new_books_found > 0:
                message = f"Added {new_books_found} book(s)!"
                self.root.after(0, self.scan_complete, message, "success")
            elif duplicates_found > 0:
                self.root.after(0, self.scan_complete, f"Found {duplicates_found} duplicate(s)", "info")
            
        except Exception as e:
            print(f"Scan error: {e}")
            self.root.after(0, self.scan_complete, "Scan error occurred", "error")
    
    def scan_complete(self, message, status_type):
        """Handle scan completion"""
        self.scan_btn.config(state="normal", text="📖 Scan Now")
        
        status_messages = {
            "success": f"✅ {message}",
            "warning": f"⚠️ {message}",
            "info": f"ℹ️ {message}",
            "error": f"❌ {message}"
        }
        
        self.status_label.config(text=status_messages.get(status_type, message))
        
        # Reset status after delay
        self.root.after(3000, lambda: self.status_label.config(text="Camera: Ready to Scan") if self.is_preview_running else None)
    
    def update_preview(self, frame):
        """Update camera preview"""
        try:
            height, width = frame.shape[:2]
            max_width, max_height = 400, 300
            
            scale = min(max_width/width, max_height/height)
            new_width = int(width * scale)
            new_height = int(height * scale)
            
            resized_frame = cv2.resize(frame, (new_width, new_height))
            rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)
            pil_image = Image.fromarray(rgb_frame)
            photo = ImageTk.PhotoImage(pil_image)
            
            self.preview_label.configure(image=photo, text="")
            self.preview_label.image = photo
            
        except Exception as e:
            print(f"Preview error: {e}")
    
    def add_code_to_tree(self, code_info):
        """Add book to the list"""
        # Only show ISBN for books, full barcode for others
        display_barcode = code_info['barcode'] if self.is_isbn(code_info['barcode']) else f"{code_info['type']}: {code_info['barcode']}"
        
        self.tree.insert('', 0, values=(  # Insert at top
            code_info['timestamp'],
            display_barcode,
            code_info['title'],
            code_info['author'],
            code_info['publisher']
        ))
    
    def update_stats(self):
        """Update statistics"""
        total = len(self.scanned_codes)
        books = len([c for c in self.scanned_codes if self.is_isbn(c['barcode'])])
        self.stats_label.config(text=f"Total Scanned: {total} | Books Found: {books}")
    
    def export_csv(self):
        """Export to CSV"""
        if not self.scanned_codes:
            messagebox.showwarning("Warning", "No books to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Book List as CSV"
        )
        
        if filename:
            try:
                df = pd.DataFrame(self.scanned_codes)
                df.to_csv(filename, index=False)
                messagebox.showinfo("Success", f"Exported {len(self.scanned_codes)} items to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def export_excel(self):
        """Export to Excel"""
        if not self.scanned_codes:
            messagebox.showwarning("Warning", "No books to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Book List as Excel"
        )
        
        if filename:
            try:
                df = pd.DataFrame(self.scanned_codes)
                df.to_excel(filename, index=False, engine='openpyxl')
                messagebox.showinfo("Success", f"Exported {len(self.scanned_codes)} items to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def clear_list(self):
        """Clear the book list"""
        if messagebox.askyesno("Confirm", "Clear all scanned books?"):
            self.scanned_codes.clear()
            self.tree.delete(*self.tree.get_children())
            self.update_stats()
    
    def run(self):
        """Start the application"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()
    
    def on_closing(self):
        """Handle app closing"""
        self.stop_camera()
        self.root.destroy()

if __name__ == "__main__":
    app = BarcodeScanner()
    app.run()
