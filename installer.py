import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Text, Scrollbar
import os
import shutil
import string
from pathlib import Path
from PIL import Image, ImageTk
import threading
import logging
import errno
import subprocess
import webbrowser
import sys

import requests
import tempfile
import win32api
import time
import hashlib
import ctypes
import winreg
import psutil
import struct
import stat  # Added for read-only handling
from typing import Optional, Dict, Any, List, Union
from dataclasses import dataclass
from enum import IntEnum
import py7zr  # Bundled for extraction
import configparser
import tkinter.font as tkfont
from concurrent.futures import ThreadPoolExecutor, as_completed
import urllib3

# Disable SSL warnings for f4se.silverlock.org (weak certificate)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Set console title for Task Manager
ctypes.windll.kernel32.SetConsoleTitleW("Fallout: London VR Installer")

# Enable DPI awareness for proper scaling on high DPI displays
try:
    # Try Windows 10+ Per-Monitor DPI awareness
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        # Fallback to Windows 8.1 System DPI awareness
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        try:
            # Fallback to Windows Vista+ DPI awareness
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass  # DPI awareness not available

# ===== INTEGRATED DOWNGRADER CLASSES =====
class ArchiveVersionEnum(IntEnum):
    """BA2 Archive version constants"""
    FALLOUT_4 = 1        # VR compatible version
    FALLOUT_4_NG2 = 7    # Next-Gen version 7
    FALLOUT_4_NG = 8     # Next-Gen version 8

@dataclass
class BA2Header:
    """BA2 Archive header structure"""
    magic: bytes           # b'BTDX'
    version: int          # Version field to modify
    format: bytes         # b'GNRL' or b'DX10'
    file_count: int       # Number of files

    @classmethod
    def from_file(cls, file_path: Union[str, Path]) -> Optional['BA2Header']:
        """Parse BA2 header from file"""
        try:
            with open(file_path, 'rb') as f:
                # Read header: magic(4) + version(4) + format(4) + file_count(4)
                header_data = f.read(16)
                if len(header_data) < 16:
                    return None
                magic = header_data[0:4]
                version = struct.unpack('<I', header_data[4:8])[0]
                format_type = header_data[8:12]
                file_count = struct.unpack('<I', header_data[12:16])[0]
                # Validate magic signature
                if magic != b'BTDX':
                    logging.warning(f"Invalid BA2 magic signature in {file_path}: {magic}")
                    return None
                return cls(magic, version, format_type, file_count)
        except Exception as e:
            logging.error(f"Error reading BA2 header from {file_path}: {e}")
            return None

class ThreadSafeCounter:
    """Thread-safe counter for tracking copied bytes"""
    
    def __init__(self, initial_value: int = 0):
        self._value = initial_value
        self._lock = threading.Lock()
    
    def add(self, amount: int) -> int:
        """Add amount and return new total"""
        with self._lock:
            self._value += amount
            return self._value
    
    def get(self) -> int:
        """Get current value"""
        with self._lock:
            return self._value
    
    def reset(self):
        """Reset to zero"""
        with self._lock:
            self._value = 0

class FalloutVRDowngrader:
    """Integrated downgrader for BA2 archives only"""

    def __init__(self, fallout4vr_data_path: Union[str, Path], progress_callback=None):
        self.data_path = Path(fallout4vr_data_path)
        self.progress_callback = progress_callback
        if not self.data_path.exists():
            raise FileNotFoundError(f"Fallout 4 VR Data directory not found: {self.data_path}")
        logging.info(f"Initialized downgrader for: {self.data_path}")

    def find_dlc_ba2_status(self) -> Dict[str, tuple[List[Path], str]]:
        """Find DLC BA2 files and their downgrade status"""
        dlc_groups = {
            "Automatron DLC": ["DLCRobot"],
            "Far Harbor DLC": ["DLCCoast"], 
            "Nuka-World DLC": ["DLCNukaWorld"],
            "Wasteland Workshop DLC": ["DLCworkshop"]
        }
        
        dlc_status = {}
        ba2_files = list(self.data_path.glob("*.ba2"))
        logging.info(f"Scanning {len(ba2_files)} BA2 files for DLC groups...")
        
        for dlc_name, prefixes in dlc_groups.items():
            dlc_files = []
            needs_downgrade = False
            
            for ba2_file in ba2_files:
                # Check if file belongs to this DLC
                if any(ba2_file.name.startswith(prefix) for prefix in prefixes):
                    dlc_files.append(ba2_file)
                    logging.info(f"Found {dlc_name} file: {ba2_file.name}")
                    header = BA2Header.from_file(ba2_file)
                    if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                        needs_downgrade = True
                        logging.info(f"  - Needs downgrade (version {header.version})")
                    elif header:
                        logging.info(f"  - London ready (version {header.version})")
            
            if dlc_files:
                status = "Needs Downgrade" if needs_downgrade else "London Ready"
                dlc_status[dlc_name] = (dlc_files, status)
                logging.info(f"{dlc_name}: {len(dlc_files)} files, status: {status}")
            else:
                logging.info(f"{dlc_name}: No files found")
            
        return dlc_status

    def get_dlc_needing_downgrade(self) -> Dict[str, List[Path]]:
        """Get only DLC groups that need downgrading"""
        dlc_status = self.find_dlc_ba2_status()
        needing_downgrade = {}
        
        for dlc_name, (files, status) in dlc_status.items():
            if status == "Needs Downgrade":
                next_gen_files = []
                for file_path in files:
                    header = BA2Header.from_file(file_path)
                    if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                        next_gen_files.append(file_path)
                if next_gen_files:
                    needing_downgrade[dlc_name] = next_gen_files
        
        return needing_downgrade

    def find_ba2_needing_downgrade(self, data_dir):
        """Find BA2 files that need downgrading"""
        next_gen_files = []
        ba2_files = list(Path(data_dir).glob("*.ba2"))
        for ba2_file in ba2_files:
            header = BA2Header.from_file(ba2_file)
            if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                next_gen_files.append(ba2_file)
        return next_gen_files

    def downgrade_ba2_file(self, file_path: Path, backup: bool = False) -> bool:
        """Downgrade a single BA2 file (no backup)"""
        try:
            header = BA2Header.from_file(file_path)
            if not header:
                logging.error(f"Could not read BA2 header from {file_path}")
                return False
            if header.version == ArchiveVersionEnum.FALLOUT_4:
                logging.info(f"BA2 file {file_path.name} already has VR compatible version")
                return True
            if header.version not in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                logging.warning(f"BA2 file {file_path.name} has unknown version {header.version}")
                return False
            
            # No backup creation
            with open(file_path, 'r+b') as f:
                f.seek(4)
                f.write(struct.pack('<I', ArchiveVersionEnum.FALLOUT_4))
            logging.info(f"Downgraded BA2: {file_path.name} ({header.version} → {ArchiveVersionEnum.FALLOUT_4})")
            return True
        except Exception as e:
            logging.error(f"Error downgrading BA2 file {file_path}: {e}")
            return False

    def downgrade_dlc_ba2_files(self) -> tuple[int, Dict[str, List[Path]]]:
        """Downgrade DLC BA2 files that need it"""
        dlc_needing_downgrade = self.get_dlc_needing_downgrade()
        
        if not dlc_needing_downgrade:
            logging.info("No DLC BA2 files found that need downgrading")
            return 0, {}
        
        success_count = 0
        downgraded_by_dlc = {}
        
        total_files = sum(len(files) for files in dlc_needing_downgrade.values())
        processed_files = 0
        
        for dlc_name, files in dlc_needing_downgrade.items():
            downgraded_files = []
            for file_path in files:
                if self.downgrade_ba2_file(file_path, backup=False):  # No backup
                    success_count += 1
                    downgraded_files.append(file_path)
                processed_files += 1
                if self.progress_callback:
                    self.progress_callback(processed_files / total_files * 100)
            
            if downgraded_files:
                downgraded_by_dlc[dlc_name] = downgraded_files
        
        logging.info(f"Successfully downgraded {success_count} DLC BA2 files")
        return success_count, downgraded_by_dlc

class ESMPatcher:
    """Apply xdelta3 patches during Fallout London VR installation"""
    def __init__(self, installer_instance):
        self.installer = installer_instance
        self.assets_dir = os.path.join(
            getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
            "assets"
        )
        self.xdelta_path = os.path.join(self.assets_dir, "xdelta3.exe")
        
        # Exact size mappings for patches
        self.patch_mappings = {
            330777465: {  # Exactly 330,777,465 bytes
                "patch": "fallout4_323025.xdelta",
                "description": "323,025 KB variant"
            },
            330553163: {  # Exactly 330,553,163 bytes
                "patch": "fallout4_322806.xdelta",
                "description": "322,806 KB variant"
            }
        }

    def patch_fallout4_esm(self, folon_data_dir):
        """Apply the appropriate ESM patch to Fallout4.esm based on its size"""
        esm_path = os.path.join(folon_data_dir, "Fallout4.esm")
        
        # Check if Fallout4.esm exists
        if not os.path.exists(esm_path):
            logging.error("Fallout4.esm not found in mod directory")
            return False
        
        # Get file size
        esm_size = os.path.getsize(esm_path)
        logging.info(f"Fallout4.esm size: {esm_size:,} bytes ({esm_size/(1024*1024):.2f} MB)")
        
        # Check for exact size match
        if esm_size not in self.patch_mappings:
            logging.warning(f"Fallout4.esm size ({esm_size:,} bytes) doesn't match known sizes: {', '.join(f'{size:,}' for size in self.patch_mappings.keys())}")
            return False
        
        # Get the matching patch
        selected_patch = self.patch_mappings[esm_size]
        patch_path = os.path.join(self.assets_dir, selected_patch["patch"])
        
        if not os.path.exists(patch_path):
            logging.error(f"Patch file not found at {patch_path}")
            return False
        
        # Check if xdelta3.exe exists
        if not os.path.exists(self.xdelta_path):
            logging.error(f"xdelta3.exe not found at {self.xdelta_path}")
            return False
        
        try:
            # Create backup
            backup_path = esm_path + '.backup'
            logging.info(f"Creating backup: {backup_path}")
            shutil.copy2(esm_path, backup_path)
            
            # Create temporary output file
            temp_output = esm_path + '.patched'
            
            # Apply patch
            cmd = [
                self.xdelta_path,
                "-f", # Force overwrite
                "-d", # Decode
                "-s", esm_path, # Source file
                patch_path, # Selected patch file
                temp_output # Output file
            ]
            
            logging.info(f"Running: {' '.join(cmd)}")
            logging.info(f"Using patch: {selected_patch['patch']} for {selected_patch['description']}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if result.returncode == 0:
                # Verify the patched file exists
                if os.path.exists(temp_output):
                    patched_size = os.path.getsize(temp_output)
                    
                    # Very lenient check - just make sure file exists and is reasonable size
                    if patched_size > 300000000: # Greater than ~286 MB
                        # Replace original with patched version
                        os.remove(esm_path)
                        os.rename(temp_output, esm_path)
                        
                        # Remove backup
                        if os.path.exists(backup_path):
                            os.remove(backup_path)
                            logging.info(f"Backup removed: {backup_path}")
                        
                        logging.info(f"✓ ESM patched successfully! New size: {patched_size:,} bytes ({patched_size/(1024*1024):.2f} MB)")
                        return True
                    else:
                        logging.error(f"Patched file seems too small: {patched_size:,} bytes")
                        # Restore backup
                        if os.path.exists(backup_path):
                            if os.path.exists(esm_path):
                                os.remove(esm_path)
                            os.rename(backup_path, esm_path)
                        if os.path.exists(temp_output):
                            os.remove(temp_output)
                        return False
                else:
                    logging.error("Patched file was not created")
                    # Restore backup
                    if os.path.exists(backup_path):
                        if os.path.exists(esm_path):
                            os.remove(esm_path)
                        os.rename(backup_path, esm_path)
                    return False
            else:
                # Patch failed - log but don't raise exception
                logging.error(f"xdelta3 failed with return code {result.returncode}")
                logging.error(f"xdelta3 stderr: {result.stderr}")
                
                # Restore backup
                if os.path.exists(backup_path):
                    if os.path.exists(esm_path):
                        os.remove(esm_path)
                    os.rename(backup_path, esm_path)
                    logging.info("Restored original Fallout4.esm from backup")
                
                # Clean up temp file
                if os.path.exists(temp_output):
                    os.remove(temp_output)
                
                return False
                
        except subprocess.TimeoutExpired:
            logging.error("Patch process timed out")
            # Restore backup
            if os.path.exists(backup_path):
                if os.path.exists(esm_path):
                    os.remove(esm_path)
                os.rename(backup_path, esm_path)
            if os.path.exists(temp_output):
                os.remove(temp_output)
            return False
            
        except Exception as e:
            logging.error(f"Error applying ESM patch: {e}")
            # Restore backup
            if os.path.exists(backup_path):
                try:
                    if os.path.exists(esm_path):
                        os.remove(esm_path)
                    os.rename(backup_path, esm_path)
                    logging.info("Restored original Fallout4.esm from backup")
                except Exception as restore_error:
                    logging.error(f"Failed to restore backup: {restore_error}")
            if os.path.exists(temp_output):
                os.remove(temp_output)
            return False

class SlideshowFrame:
    """Slideshow frame to display installation images"""
   
    def __init__(self, parent, installer_instance):
        self.installer = installer_instance
        self.parent = parent
        self.current_image_index = 0
        self.slideshow_running = False
        self.slideshow_after_id = None
       
        # Store references to PhotoImage objects to prevent garbage collection
        self.photo_images = []
       
        logging.info("=== SLIDESHOW INITIALIZATION STARTED ===")
        logging.info(f"Parent widget type: {type(parent)}")
        logging.info(f"Parent widget exists: {parent.winfo_exists()}")
       
        # Create slideshow frame with better visibility settings
        self.slideshow_frame = tk.Frame(parent, bg="#1e1e1e")
        logging.info(f"Slideshow frame created: {self.slideshow_frame}")
       
        # Create a container frame for better layout control
        self.container_frame = tk.Frame(self.slideshow_frame, bg="#1e1e1e")
        self.container_frame.pack(expand=False, padx=self.installer.get_scaled_value(20), pady=self.installer.get_scaled_value(20))
       
        # Create image label with border for debugging
        self.image_label = tk.Label(
            self.container_frame,
            bg="#2e2e2e",  # Slightly lighter background for visibility
            borderwidth=0,
            relief="flat"
        )
        self.image_label.pack()
        logging.info(f"Image label created: {self.image_label}")
       
        # Load all images
        self.load_slideshow_images()
   
    def load_slideshow_images(self):
        """Load all slideshow images into memory"""
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
        assets_dir = os.path.join(base_dir, "assets")
        
        # Fallback: if assets_dir doesn't exist, check if files are in base_dir directly
        if not os.path.exists(assets_dir):
            # Check if PNG files exist in base_dir directly
            if os.path.exists(os.path.join(base_dir, "installerLS1.png")):
                assets_dir = base_dir
       
        logging.info(f"=== LOADING SLIDESHOW IMAGES ===")
        logging.info(f"Assets directory: {assets_dir}")
        logging.info(f"Assets directory exists: {os.path.exists(assets_dir)}")
       
        # List all files in assets directory
        if os.path.exists(assets_dir):
            all_files = os.listdir(assets_dir)
            png_files = [f for f in all_files if f.lower().endswith('.png')]
            logging.info(f"Total files in assets: {len(all_files)}")
            logging.info(f"PNG files found: {png_files}")
        else:
            logging.error(f"Assets directory does not exist: {assets_dir}")
            return
       
        self.images = []
        self.photo_images = []  # Keep references to prevent garbage collection
       
        # Try to load images installerLS1.png through installerLS12.png
        for i in range(1, 13):
            image_filename = f"installerLS{i}.png"
            image_path = os.path.join(assets_dir, image_filename)
           
            logging.info(f"Attempting to load: {image_filename}")
            logging.info(f"Full path: {image_path}")
            logging.info(f"File exists: {os.path.exists(image_path)}")
           
            if os.path.exists(image_path):
                try:
                    # Get file size for debugging
                    file_size = os.path.getsize(image_path)
                    logging.info(f"File size: {file_size} bytes")
                   
                    # Load and resize image
                    img = Image.open(image_path)
                    logging.info(f"Image loaded successfully: {image_filename}")
                    logging.info(f"Original size: {img.size}, Mode: {img.mode}")
                   
                    # For PNGs with transparency, composite onto the dark background color
                    if img.mode in ('RGBA', 'LA'):
                        # Create background matching the installer's dark theme
                        background = Image.new('RGBA', img.size, (30, 30, 30, 255))  # #1e1e1e
                        background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                        img = background.convert('RGB')
                    elif img.mode not in ('RGB', 'L'):
                        img = img.convert('RGB')
                   
                    # Calculate size to fit within window while maintaining aspect ratio
                    max_width = self.installer.get_scaled_value(500)
                    max_height = self.installer.get_scaled_value(700)
                   
                    # Get original dimensions
                    orig_width, orig_height = img.size
                   
                    # Calculate scaling factor
                    scale = min(max_width/orig_width, max_height/orig_height)
                   
                    # Calculate new dimensions
                    new_width = int(orig_width * scale)
                    new_height = int(orig_height * scale)
                   
                    logging.info(f"Resizing to: {new_width}x{new_height}")
                   
                    # Resize image
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                   
                    # Convert to PhotoImage and store
                    photo = ImageTk.PhotoImage(img)
                    self.photo_images.append(photo)
                    self.images.append({
                        'photo': photo,
                        'name': f"installerLS{i}",
                        'path': image_path
                    })
                   
                    logging.info(f"✓ Successfully loaded and processed: {image_filename}")
                   
                except Exception as e:
                    logging.error(f"✗ Failed to load slideshow image {image_filename}: {e}")
                    import traceback
                    logging.error(traceback.format_exc())
            else:
                logging.warning(f"✗ Slideshow image not found: {image_path}")
       
        logging.info(f"=== SLIDESHOW LOADING COMPLETE ===")
        logging.info(f"Total images loaded: {len(self.images)}")
       
        if not self.images:
            logging.error("NO SLIDESHOW IMAGES WERE LOADED!")
            # Show error message in UI (optional, since caption is removed)
   
    def start_slideshow(self):
        """Start the slideshow"""
        logging.info("=== STARTING SLIDESHOW ===")
        logging.info(f"Images available: {len(self.images)}")
       
        if not self.images:
            logging.error("Cannot start slideshow: No images available")
            return
       
        self.slideshow_running = True
        self.current_image_index = 0
       
        # Pack the slideshow frame centered with space from progress bar
        self.slideshow_frame.pack(anchor="center", pady=(self.installer.get_scaled_value(10), self.installer.get_scaled_value(10)))
        logging.info(f"Slideshow frame packed")
        logging.info(f"Frame visible: {self.slideshow_frame.winfo_viewable()}")
        logging.info(f"Frame dimensions: {self.slideshow_frame.winfo_width()}x{self.slideshow_frame.winfo_height()}")
       
        # Force update to ensure frame is displayed
        self.slideshow_frame.update_idletasks()
       
        # Display first image immediately
        self.update_slideshow_image()
       
        # Schedule next image change
        self.schedule_next_image()
       
        logging.info("Slideshow started successfully")
   
    def update_slideshow_image(self):
        """Update the displayed image"""
        if not self.slideshow_running or not self.images:
            logging.warning("Cannot update image: slideshow not running or no images")
            return
       
        try:
            current_image = self.images[self.current_image_index]
            logging.info(f"Displaying image {self.current_image_index + 1}/{len(self.images)}: {current_image['name']}")
           
            # Update image
            self.image_label.config(image=current_image['photo'])
            
            # Apply per-image positioning adjustments for off-center images
            # LS5 and LS6 need to be shifted left to appear centered
            if current_image['name'] in ['installerLS5', 'installerLS6']:
                self.image_label.pack_configure(padx=(0, self.installer.get_scaled_value(40)))  # Shift left
            elif current_image['name'] == 'installerLS12':
                self.image_label.pack_configure(padx=(self.installer.get_scaled_value(40), 0))  # Shift right
            else:
                self.image_label.pack_configure(padx=0)  # Default centered
           
            # Force the label to update
            self.image_label.update_idletasks()
           
            # Log widget states
            logging.info(f"Image label visible: {self.image_label.winfo_viewable()}")
            logging.info(f"Image label size: {self.image_label.winfo_width()}x{self.image_label.winfo_height()}")
           
        except Exception as e:
            logging.error(f"Error updating slideshow image: {e}")
            import traceback
            logging.error(traceback.format_exc())
   
    def schedule_next_image(self):
        """Schedule the next image change"""
        if self.slideshow_running:
            # Cancel any existing scheduled change
            if self.slideshow_after_id:
                try:
                    self.parent.after_cancel(self.slideshow_after_id)
                except:
                    pass
           
            # Schedule next change in 20 seconds (20000 ms)
            self.slideshow_after_id = self.parent.after(20000, self.next_image)
            logging.info(f"Next image scheduled in 20 seconds")
   
    def next_image(self):
        """Move to the next image in the slideshow"""
        if not self.slideshow_running or not self.images:
            return
       
        # Move to next image (loop back to start if at end)
        self.current_image_index = (self.current_image_index + 1) % len(self.images)
        logging.info(f"Advancing to image {self.current_image_index + 1}/{len(self.images)}")
       
        # Update display
        self.update_slideshow_image()
       
        # Schedule next change
        self.schedule_next_image()
   
    def stop_slideshow(self):
        """Stop the slideshow"""
        logging.info("=== STOPPING SLIDESHOW ===")
        self.slideshow_running = False
       
        # Cancel scheduled image change
        if self.slideshow_after_id:
            try:
                self.parent.after_cancel(self.slideshow_after_id)
                self.slideshow_after_id = None
            except:
                pass
       
        # Hide slideshow frame
        self.slideshow_frame.pack_forget()
       
        logging.info("Slideshow stopped")
   
    def cleanup(self):
        """Clean up resources"""
        self.stop_slideshow()
        self.photo_images.clear()
        self.images.clear()
        logging.info("Slideshow resources cleaned up")

class FalloutLondonVRInstaller:
    def __init__(self, root, skip_update_detection=False, initial_install_path=None):
        self.root = root
        self.root.title("")
        self.skip_update_detection = skip_update_detection  # Flag to skip update mode detection
        self.initial_install_path = initial_install_path  # Preserved install path from restart
        
        # Detect screen resolution and calculate scaling
        self.setup_dynamic_sizing()
        
        self.root.configure(bg="#1e1e1e")

        # Setup fonts with scaling
        self.setup_scaled_fonts()

        # Setup window only once
        self.root.withdraw()  # Hide window temporarily
        self.root.update_idletasks()  # Process pending events
        self.setup_borderless_with_taskbar_immediate()  # Set up immediately
        
        # Set custom icon
        self.setup_custom_icon()

        self.f4_path = tk.StringVar()
        self.f4vr_path = tk.StringVar()
        self.london_data_path = tk.StringVar()
        self.mo2_path = tk.StringVar()
        self.existing_install_path = tk.StringVar()
        self.london_installed = False
        self.f4_version = "Unknown"
        self.dlc_status = {}
        self.missing_dlc = []
        self.needs_downgrade = False
        self.cancel_requested = False
        self.update_mode = False
        self.is_update_detected = False
        self.detected_install_path = None
        self.installation_mode = tk.StringVar(value="fresh")
        self.progress = None
        self.progress_label = None
        self.message_label = None
        self.london_data_entry = None
        self.browse_button = None
        self.install_button = None
        self.nexus_username = tk.StringVar()
        self.nexus_password = tk.StringVar()
        self.ini_updated = False
        self.path_display_text = None
        self.path_display_label = None
        self.location_label = None
        self.slideshow = None  # Add with other instance variables
        self.xse_preloader_installed = False  # Track xSE Plugin Preloader installation
        self.welcome_canvas = None  # For scrollable welcome page
        self.upgrade_to_103 = False  # Flag for upgrading from 1.02 to 1.03
        self.london_103_source_path = None  # Path to London 1.03 files for upgrade

        # Load logo with better error handling
        try:
            logo_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "logo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_size = self.get_scaled_value(115)  # 15% larger than 100
                logo_img = logo_img.resize((logo_size, logo_size), Image.Resampling.LANCZOS)
                self.logo = ImageTk.PhotoImage(logo_img)
            else:
                self.logo = None
                logging.warning(f"Logo file not found at {logo_path}")
        except Exception as e:
            self.logo = None
            logging.error(f"Failed to load logo: {e}")

        # Load background image
        try:
            bg_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "background.png")
            if os.path.exists(bg_path):
                bg_img = Image.open(bg_path)
                bg_img = bg_img.resize((self.window_width, self.window_height), Image.Resampling.LANCZOS)  # Use scaled dimensions
                self.bg_image = ImageTk.PhotoImage(bg_img)
            else:
                self.bg_image = None
                logging.warning(f"Background image not found at {bg_path}")
        except Exception as e:
            self.bg_image = None
            logging.error(f"Failed to load background image: {e}")

        # Load atkins image
        try:
            atkins_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "atkins.png")
            logging.info(f"Attempting to load atkins image from: {atkins_path}")
            if os.path.exists(atkins_path):
                logging.info("Atkins image file found, loading...")
                atkins_img = Image.open(atkins_path)
                logging.info(f"Atkins image opened successfully, size: {atkins_img.width}x{atkins_img.height}")
                # Scale the image to use maximum available space
                atkins_width = self.get_scaled_value(310)  # Reduced 10% from 345
                aspect_ratio = atkins_img.height / atkins_img.width
                atkins_height = int(atkins_width * aspect_ratio)
                atkins_img = atkins_img.resize((atkins_width, atkins_height), Image.Resampling.LANCZOS)
                self.atkins_image = ImageTk.PhotoImage(atkins_img)
                logging.info(f"Atkins image loaded successfully, scaled to {atkins_width}x{atkins_height}")
            else:
                self.atkins_image = None
                logging.warning(f"Atkins image not found at {atkins_path}")
        except Exception as e:
            self.atkins_image = None
            logging.error(f"Failed to load atkins image: {e}")

        # Setup logging with proper encoding - use exe directory or temp for compiled
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            log_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            log_dir = os.path.dirname(__file__)
        log_path = os.path.join(log_dir, "installer.log")
        
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler(log_path, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )

        # Create the welcome page directly (skip mode selection)
        self.create_welcome_page()
        
        # Initialize slideshow
        self.slideshow = None
        
        # Center window once
        self.center_window()
        
        # Note: Window will be shown AFTER detect_paths() completes to avoid flicker
        # when switching to update mode. See end of detect_paths() for deiconify() call.

    def setup_dynamic_sizing(self):
        """Detect screen resolution and calculate appropriate window size and scaling"""
        try:
            # Get screen dimensions (these are in virtual/scaled pixels on high-DPI displays)
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            logging.info(f"Detected screen resolution: {screen_width}x{screen_height}")
            
            # Calculate usable area (leave space for taskbar and window decorations)
            usable_width = screen_width * 0.9  # Use 90% of screen width
            usable_height = screen_height * 0.85  # Use 85% of screen height (account for taskbar)
            
            # Base dimensions (original design)
            base_width = 553
            base_height = 990
            
            # Calculate scaling factors based on available screen space
            width_scale = usable_width / base_width
            height_scale = usable_height / base_height
            
            # Use the smaller scale to maintain aspect ratio
            self.ui_scale = min(width_scale, height_scale)
            
            # Cap scaling - don't scale up beyond 1.0 (let OS handle high-DPI scaling)
            # and don't scale down too much for small screens
            self.ui_scale = max(0.55, min(self.ui_scale, 1.0))
            
            # Calculate final window dimensions
            self.window_width = int(base_width * self.ui_scale)
            self.window_height = int(base_height * self.ui_scale)
            
            # Ensure minimum usable size
            self.window_width = max(self.window_width, 400)
            self.window_height = max(self.window_height, 500)
            
            # Ensure it fits on screen with some margin
            self.window_width = min(self.window_width, int(screen_width - 20))
            self.window_height = min(self.window_height, int(screen_height - 60))  # Account for taskbar
            
            logging.info(f"UI scale factor: {self.ui_scale:.2f}")
            logging.info(f"Final window size: {self.window_width}x{self.window_height}")
            
            # Set window geometry - allow resizing for small screens
            self.root.geometry(f"{self.window_width}x{self.window_height}")
            # Set a smaller minimum size to allow users to resize if needed
            self.root.minsize(400, 450)
            # Don't set maxsize - allow resizing on small/high-DPI screens
            
        except Exception as e:
            logging.warning(f"Failed to setup dynamic sizing, using defaults: {e}")
            # Fallback to original dimensions
            self.ui_scale = 1.0
            self.window_width = 553
            self.window_height = 990
            self.root.geometry(f"{self.window_width}x{self.window_height}")
            self.root.minsize(400, 450)

    def setup_scaled_fonts(self):
        """Setup fonts that scale with UI"""
        try:
            # Base font sizes
            base_regular = 12
            base_title = 16
            
            # Scale font sizes
            regular_size = max(8, int(base_regular * self.ui_scale))
            title_size = max(10, int(base_title * self.ui_scale))
            
            # Create scaled fonts
            self.regular_font = ("Segoe UI", regular_size)
            self.bold_font = ("Segoe UI", regular_size, "bold")
            self.title_font = ("Segoe UI", title_size, "bold")
            
            logging.info(f"Scaled fonts - Regular: {regular_size}pt, Title: {title_size}pt")
            
        except Exception as e:
            logging.warning(f"Failed to setup scaled fonts, using defaults: {e}")
            # Fallback fonts
            self.regular_font = ("Segoe UI", 12)
            self.bold_font = ("Segoe UI", 12, "bold")
            self.title_font = ("Segoe UI", 16, "bold")

    def get_scaled_value(self, base_value):
        """Get a scaled value based on UI scale factor"""
        return max(1, int(base_value * self.ui_scale))
    
    def get_scaled_padding(self, base_padding):
        """Get scaled padding values - returns tuple for consistency"""
        if isinstance(base_padding, (list, tuple)):
            return tuple(self.get_scaled_value(p) for p in base_padding)
        return self.get_scaled_value(base_padding)

    def center_window(self):
        """Center window on screen - call only once"""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Calculate centered position
        x = (screen_width // 2) - (self.window_width // 2)
        y = (screen_height // 2) - (self.window_height // 2)
        
        # Ensure window stays within screen bounds
        x = max(0, min(x, screen_width - self.window_width))
        y = max(0, min(y, screen_height - self.window_height))
        
        self.root.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")

    def setup_window(self):
        """Setup window position and properties"""
        try:
            # Center window on screen
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Calculate centered position
            x = (screen_width // 2) - (width // 2)
            y = (screen_height // 2) - (height // 2)
            
            # Ensure window stays within screen bounds
            x = max(0, min(x, screen_width - width))
            y = max(0, min(y, screen_height - height))
            
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            
            # Use smaller minimum to support high-DPI small screens
            self.root.minsize(400, 450)
        except Exception as e:
            logging.warning(f"Could not setup window: {e}")

    def setup_borderless_with_taskbar(self):
        """Remove title bar but keep taskbar presence"""
        try:
            from ctypes import windll
            
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            
            # Get window handle
            hwnd = windll.user32.GetParent(self.root.winfo_id())
            
            # Get current style
            style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            
            # Remove tool window, add app window
            style = style & ~WS_EX_TOOLWINDOW
            style = style | WS_EX_APPWINDOW
            
            # Apply new style
            windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            
            # Remove window frame
            self.root.overrideredirect(True)
            
            # Refresh window
            self.root.withdraw()
            self.root.after(10, lambda: self.root.deiconify())
            
            logging.info("Set borderless window with taskbar presence")
            
        except Exception as e:
            logging.warning(f"Could not set borderless window: {e}")

    def setup_borderless_with_taskbar_immediate(self):
        """Remove title bar but keep taskbar presence - immediate version"""
        try:
            from ctypes import windll
            
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            
            # Remove window frame first
            self.root.overrideredirect(True)
            
            # Force window to update
            self.root.update_idletasks()
            
            # Get window handle
            hwnd = windll.user32.GetParent(self.root.winfo_id())
            
            # Get current style
            style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            
            # Remove tool window, add app window
            style = style & ~WS_EX_TOOLWINDOW
            style = style | WS_EX_APPWINDOW
            
            # Apply new style
            windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            
            # Force window to refresh its frame
            windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 
                                     0x0001 | 0x0002 | 0x0004 | 0x0020)
            
            logging.info("Set borderless window with taskbar presence (immediate)")
            
        except Exception as e:
            logging.warning(f"Could not set borderless window: {e}")
            # Fall back to normal window if it fails
            self.root.overrideredirect(False)

    def setup_custom_icon(self):
        """Set custom icon for both title bar and taskbar"""
        try:
            icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                logging.info(f"Set custom icon from {icon_path}")
            else:
                logging.warning(f"Icon not found at {icon_path}")
        except Exception as e:
            logging.warning(f"Could not set custom icon: {e}")

    def add_custom_title_bar(self):
        """Add custom title bar with controls"""
        try:
            logging.debug("Starting title bar creation for root window")
            title_bar = tk.Frame(self.root, bg="#2d2d2d", height=30)
            title_bar.pack(fill="x", side="top")
            title_bar.pack_propagate(False)
            
            title_label = tk.Label(title_bar, text="", 
                                  bg="#2d2d2d", fg="#ffffff", font=("Segoe UI", 10))
            title_label.pack(side="left", padx=10, pady=5)
            
            close_btn = tk.Button(title_bar, text="✕", command=lambda: sys.exit(0),
                                 bg="#2d2d2d", fg="#ffffff", bd=0, font=("Segoe UI", 12),
                                 activebackground="#ff4444", width=3)
            close_btn.pack(side="right", padx=5, pady=2)
            
            minimize_btn = tk.Button(title_bar, text="—", command=lambda: self.minimize_window(self.root),
                                    bg="#2d2d2d", fg="#ffffff", bd=0, font=("Segoe UI", 12),
                                    activebackground="#444444", width=3, cursor="hand2")
            minimize_btn.pack(side="right", padx=2, pady=2)
            
            self.make_draggable(title_bar)
            logging.debug("Title bar creation completed for root window")
        except Exception as e:
            logging.error(f"Failed to create custom title bar for root: {e}")

    def add_custom_title_bar_to_window(self, window):
        """Add custom title bar to a specified window (root or Toplevel)"""
        try:
            logging.debug("Starting title bar creation for window")
            title_bar = tk.Frame(window, bg="#2d2d2d", height=30)
            title_bar.pack(fill="x", side="top")
            title_bar.pack_propagate(False)
            
            title_label = tk.Label(title_bar, text="", 
                                  bg="#2d2d2d", fg="#ffffff", font=("Segoe UI", 10))
            title_label.pack(side="left", padx=10, pady=5)
            
            close_btn = tk.Button(title_bar, text="✕", command=window.destroy,
                                 bg="#2d2d2d", fg="#ffffff", bd=0, font=("Segoe UI", 12),
                                 activebackground="#ff4444", width=3)
            close_btn.pack(side="right", padx=5, pady=2)
            
            minimize_btn = tk.Button(title_bar, text="—", command=lambda: self.minimize_window(window),
                                    bg="#2d2d2d", fg="#ffffff", bd=0, font=("Segoe UI", 12),
                                    activebackground="#444444", width=3, cursor="hand2")
            minimize_btn.pack(side="right", padx=2, pady=2)
            
            self.make_draggable(title_bar)
            logging.debug("Title bar creation completed for window")
        except Exception as e:
            logging.error(f"Failed to create custom title bar: {e}")

    def make_draggable(self, widget):
        """Make widget draggable to move window"""
        def start_drag(event):
            widget.start_x = event.x
            widget.start_y = event.y
        
        def drag_window(event):
            x = self.root.winfo_x() + (event.x - widget.start_x)
            y = self.root.winfo_y() + (event.y - widget.start_y)
            self.root.geometry(f"+{x}+{y}")
        
        widget.bind("<Button-1>", start_drag)
        widget.bind("<B1-Motion>", drag_window)

    def minimize_window(self, window):
        """Minimize the specified window using Windows API"""
        try:
            from ctypes import windll
            SW_MINIMIZE = 6  # Windows API constant for minimizing
            hwnd = windll.user32.GetParent(window.winfo_id())
            windll.user32.ShowWindow(hwnd, SW_MINIMIZE)
            logging.info(f"Minimized window: {window}")
        except Exception as e:
            logging.error(f"Failed to minimize window: {e}")

    def sanitize_path(self, path):
        """Sanitize and validate file paths"""
        if not path:
            raise ValueError("Path cannot be empty")
        
        # Normalize path
        path = os.path.normpath(path)
        
        # Check for invalid characters (excluding colon which is valid for drive letters)
        invalid_chars = ['<', '>', '"', '|', '?', '*']
        
        # Special handling for colon - only invalid if not part of drive letter
        if ':' in path:
            # Check if colon is in valid position (drive letter)
            colon_positions = [i for i, char in enumerate(path) if char == ':']
            for pos in colon_positions:
                # Valid colon positions: position 1 for drive letters (C:)
                if pos != 1:
                    raise ValueError(f"Invalid character ':' in path at position {pos}")
        
        # Check other invalid characters
        for char in invalid_chars:
            if char in path:
                raise ValueError(f"Invalid character '{char}' in path")
        
        return path

    def safe_widget_update(self, widget, update_func, *args, **kwargs):
        """Safely update a widget with existence checking"""
        if widget and hasattr(widget, 'winfo_exists'):
            try:
                if widget.winfo_exists():
                    self.root.after(0, lambda: update_func(*args, **kwargs) if widget.winfo_exists() else None)
            except tk.TclError:
                logging.debug(f"Widget no longer exists: {widget}")

    def update_message(self, text, color="#ffffff"):
        """Safely update message label"""
        if hasattr(self, 'message_label') and self.message_label:
            try:
                if self.message_label.winfo_exists():
                    self.root.after(0, lambda t=text, c=color: 
                        self.message_label.config(text=t, fg=c) 
                        if self.message_label.winfo_exists() else None)
            except tk.TclError:
                logging.debug("Message label no longer exists")

    def update_progress(self, value, label_text=None):
        """Safely update progress bar and label"""
        if hasattr(self, 'progress') and self.progress:
            try:
                if self.progress.winfo_exists():
                    self.root.after(0, lambda v=value: 
                        self.progress.__setitem__("value", v) 
                        if self.progress.winfo_exists() else None)
            except tk.TclError:
                logging.debug("Progress bar no longer exists")
        
        if label_text and hasattr(self, 'progress_label') and self.progress_label:
            try:
                if self.progress_label.winfo_exists():
                    self.root.after(0, lambda lt=label_text: 
                        self.progress_label.config(text=lt) 
                        if self.progress_label.winfo_exists() else None)
            except tk.TclError:
                logging.debug("Progress label no longer exists")

    def perform_installation(self):
        """Perform the complete installation process"""
        try:
            logging.info("=== INSTALLATION STARTED ===")
            logging.info(f"User selected F4VR path: {self.f4vr_path.get()}")
            logging.info(f"User selected MO2 path: {self.mo2_path.get()}")
            logging.info(f"User selected F4 path: {self.f4_path.get()}")
            logging.info(f"User selected London path: {self.london_data_path.get()}")
            
            if not self.mo2_path.get():
                self.root.after(0, lambda: self.message_label.config(text="Please select a valid installation directory.", fg="#ff0000") if self.message_label.winfo_exists() else None)
                return

            # Step 1: Download and install MO2
            mo2_archive = self.download_mo2_portable()
            self.extract_mo2(mo2_archive)

            # Step 2: Download and install F4SEVR to Fallout 4 VR dir
            f4sevr_archive = self.download_f4sevr()
            self.extract_and_install_f4sevr(f4sevr_archive, self.f4vr_path.get())

            # Step 3: Download and install FRIK
            self.download_and_install_frik()

            # Step 4: Download and install Comfort Swim VR
            self.download_and_install_comfort_swim()

            # Step 5: Download and install Buffout 4 NG
            self.download_and_install_buffout4()

            # Step 6: Copy MO2 assets
            self.copy_mo2_assets()
            
            # Step 7: Copy Fallout data
            self.copy_fallout_data_with_smart_dlc()
            
            # Step 8: Copy FRIK configuration
            self.copy_frik_ini()
            
            # Step 9: Install xSE Plugin Preloader
            self.install_xse_plugin_preloader()
            
            # Step 11: Apply downgrader if needed
            if self.needs_downgrade:
                self.perform_downgrade_step()
            
            # Step 12: Complete installation
            self.root.after(0, lambda: self.message_label.config(text="Installation completed successfully!", fg="#00ff00") if self.message_label.winfo_exists() else None)
            logging.info("Installation completed successfully")
            
            # Show completion buttons
            self.show_completion_ui()
            
            # Create desktop shortcut
            self.create_desktop_shortcut()
            
            # Create Start Menu shortcuts
            self.create_start_menu_shortcuts()
        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Installation failed: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            logging.error(f"Installation failed: {e}")
            raise

    def detect_steam_paths(self) -> list[str]:
        """Detect Steam installation paths from registry and libraryfolders.vdf"""
        steam_paths = []
        try:
            # Get main Steam installation path from registry
            steam_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Valve\Steam")
            steam_path = winreg.QueryValueEx(steam_key, "InstallPath")[0]
            winreg.CloseKey(steam_key)
            steam_paths.append(steam_path)
            logging.info(f"Detected Steam path from registry: {steam_path}")
            
            # Parse libraryfolders.vdf to find additional Steam library locations
            library_file = os.path.join(steam_path, "steamapps", "libraryfolders.vdf")
            if os.path.exists(library_file):
                try:
                    with open(library_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Parse VDF format - look for "path" entries
                    import re
                    # Match patterns like "path"		"D:\\SteamLibrary"
                    path_matches = re.findall(r'"path"\s*"([^"]+)"', content)
                    for lib_path in path_matches:
                        # Unescape backslashes
                        lib_path = lib_path.replace('\\\\', '\\')
                        if lib_path not in steam_paths and os.path.exists(lib_path):
                            steam_paths.append(lib_path)
                            logging.info(f"Detected Steam library from libraryfolders.vdf: {lib_path}")
                except Exception as e:
                    logging.warning(f"Failed to parse libraryfolders.vdf: {e}")
                    
        except Exception as e:
            logging.warning(f"Failed to detect Steam paths from registry: {e}")
        return steam_paths

    def detect_gog_paths(self) -> list[str]:
        """Detect GOG Galaxy installation paths"""
        gog_paths = []
        try:
            # Check GOG Galaxy registry entries
            try:
                gog_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\GOG.com\GalaxyClient\paths")
                gog_path = winreg.QueryValueEx(gog_key, "client")[0]
                winreg.CloseKey(gog_key)
                # GOG games are typically in a Games subfolder
                gog_games_path = os.path.join(os.path.dirname(gog_path), "Games")
                if os.path.exists(gog_games_path):
                    gog_paths.append(gog_games_path)
                    logging.info(f"Detected GOG path: {gog_games_path}")
            except Exception:
                pass
            
            # Check alternative GOG registry location
            try:
                gog_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\GOG.com\Games\1998527297")  # Fallout 4 GOG ID
                gog_fallout_path = winreg.QueryValueEx(gog_key, "path")[0]
                winreg.CloseKey(gog_key)
                gog_paths.append(os.path.dirname(gog_fallout_path))
                logging.info(f"Detected GOG Fallout 4 path: {gog_fallout_path}")
            except Exception:
                pass
                
        except Exception as e:
            logging.error(f"Failed to detect GOG paths: {e}")
        
        return gog_paths

    def get_fallout4_version(self, fallout4_exe_path):
        """Get version information from Fallout4.exe"""
        try:
            # Get file version info using Windows API
            version_info = win32api.GetFileVersionInfo(fallout4_exe_path, "\\")
            ms = version_info['FileVersionMS']
            ls = version_info['FileVersionLS']
            
            version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            
            # Determine if it's Next-Gen (1.10.980+) or older
            major, minor, build, revision = map(int, version.split('.'))
            
            # Next-Gen versions are 1.10.980 and higher
            is_next_gen = (major > 1) or (major == 1 and minor > 10) or (major == 1 and minor == 10 and build >= 980)
            
            return version, is_next_gen
            
        except Exception as e:
            logging.error(f"Failed to get Fallout 4 version: {e}")
            return "Unknown", False

    def verify_file_integrity(self, file_path: str, expected_hash: Optional[str] = None) -> bool:
        if not os.path.exists(file_path):
            return False
        if expected_hash:
            with open(file_path, 'rb') as f:
                file_hash = hashlib.sha256(f.read()).hexdigest()
            return file_hash == expected_hash
        return os.path.getsize(file_path) > 0

    def remove_readonly_and_overwrite(self, file_path):
        """Remove read-only attribute from file and prepare for overwriting"""
        try:
            if os.path.exists(file_path):
                # Remove read-only attribute
                current_permissions = os.stat(file_path).st_mode
                os.chmod(file_path, current_permissions | stat.S_IWRITE)
                logging.debug(f"Removed read-only attribute from {file_path}")
        except Exception as e:
            logging.warning(f"Could not remove read-only attribute from {file_path}: {e}")
            # Try alternative method for Windows
            try:
                subprocess.run(['attrib', '-R', file_path], check=False, capture_output=True)
                logging.debug(f"Removed read-only using attrib command: {file_path}")
            except Exception as attrib_error:
                logging.warning(f"Attrib command also failed for {file_path}: {attrib_error}")

    def detect_existing_installation(self):
        """Detect existing Fallout London VR installation"""
        search_paths = []
        
        # Get all active drives
        active_drives = self.get_all_active_drives()
        
        # Build search paths with exact names (with and without spaces)
        for drive in active_drives:
            search_paths.extend([
                f"{drive}\\Games\\Fallout London VR",
                f"{drive}\\Games\\FalloutLondonVR",
                f"{drive}\\Fallout London VR",
                f"{drive}\\FalloutLondonVR",
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common\\Fallout London VR",
                f"{drive}\\Steam\\steamapps\\common\\Fallout London VR",
                f"{drive}\\SteamLibrary\\steamapps\\common\\Fallout London VR",
                f"{drive}\\GOG Games\\Fallout London VR",
                f"{drive}\\Program Files (x86)\\GOG Galaxy\\Games\\Fallout London VR",
            ])
        
        # Check each exact path for valid installation
        for path in search_paths:
            if self.is_valid_existing_installation(path):
                return path
        
        # If not found, scan common parent directories for folders containing "Fallout London VR"
        common_parent_dirs = []
        for drive in active_drives:
            common_parent_dirs.extend([
                f"{drive}\\Games",
                f"{drive}\\",
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common",
                f"{drive}\\Steam\\steamapps\\common",
                f"{drive}\\SteamLibrary\\steamapps\\common",
                f"{drive}\\GOG Games",
                f"{drive}\\Program Files (x86)\\GOG Galaxy\\Games",
            ])
        
        # Scan parent directories for any folder starting with "fallout"
        for parent_dir in common_parent_dirs:
            if os.path.exists(parent_dir):
                try:
                    for item in os.listdir(parent_dir):
                        # Check any folder starting with "fallout"
                        if item.lower().startswith("fallout"):
                            full_path = os.path.join(parent_dir, item)
                            if os.path.isdir(full_path) and self.is_valid_existing_installation(full_path):
                                logging.info(f"Found installation: {full_path}")
                                return full_path
                except Exception as e:
                    logging.debug(f"Error scanning {parent_dir}: {e}")
        
        return None

    def is_valid_existing_installation(self, path):
        """Check if path contains a valid Fallout London VR installation"""
        try:
            # Check for ModOrganizer.exe
            mo2_exe = os.path.join(path, "ModOrganizer.exe")
            if not os.path.exists(mo2_exe):
                return False
            
            # Check for Fallout London VR mod in mods directory
            folon_mod_path = os.path.join(path, "mods", "Fallout London VR")
            if not os.path.exists(folon_mod_path):
                return False
            
            # Check for Fallout London VR.esp in the mod directory
            folon_esp = os.path.join(folon_mod_path, "Fallout London VR.esp")
            if os.path.exists(folon_esp):
                logging.info(f"Valid installation found at {path}")
                return True
            
            return False
        except Exception as e:
            logging.debug(f"Error checking path {path}: {e}")
            return False

    def get_installed_london_version(self, install_path):
        """
        Detect the installed Fallout: London version in an existing installation.
        Returns: "1.03", "1.02", or None if cannot determine
        """
        try:
            # Check the Fallout London Data mod folder
            london_data_path = os.path.join(install_path, "mods", "Fallout London Data")
            
            if not os.path.exists(london_data_path):
                logging.info(f"Fallout London Data folder not found at {london_data_path}")
                return None
            
            # Check for Textures14.ba2 which indicates v1.03
            textures14 = os.path.join(london_data_path, "LondonWorldSpace - Textures14.ba2")
            if os.path.exists(textures14):
                logging.info(f"Detected installed London version: 1.03 (Textures14 found)")
                return "1.03"
            
            # Check if basic London files exist (indicates 1.02)
            esm_file = os.path.join(london_data_path, "LondonWorldSpace.esm")
            if os.path.exists(esm_file):
                logging.info(f"Detected installed London version: 1.02 (no Textures14)")
                return "1.02"
            
            logging.info(f"Could not determine installed London version")
            return None
            
        except Exception as e:
            logging.error(f"Error detecting installed London version: {e}")
            return None

    def scan_for_london_103_files(self):
        """
        Scan the system for Fallout: London 1.03 files.
        Returns: path to 1.03 files if found, None otherwise
        """
        try:
            active_drives = self.get_all_active_drives()
            
            # Common locations where London files might be
            search_paths = []
            for drive in active_drives:
                search_paths.extend([
                    f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common\\Fallout 4",
                    f"{drive}\\Steam\\steamapps\\common\\Fallout 4",
                    f"{drive}\\SteamLibrary\\steamapps\\common\\Fallout 4",
                    f"{drive}\\GOG Games\\Fallout 4",
                    f"{drive}\\Program Files (x86)\\GOG Galaxy\\Games\\Fallout 4",
                    f"{drive}\\Games\\Fallout 4",
                    f"{drive}\\Fallout 4",
                    # Also check for standalone London downloads
                    f"{drive}\\Downloads",
                    f"{drive}\\Users\\{os.getlogin()}\\Downloads",
                    f"{drive}\\Fallout London",
                    f"{drive}\\Games\\Fallout London",
                ])
            
            for base_path in search_paths:
                if not os.path.exists(base_path):
                    continue
                
                # Check root
                is_valid, version, _, _ = self.validate_london_files(base_path)
                if is_valid and version == "1.03":
                    logging.info(f"Found London 1.03 files at: {base_path}")
                    return base_path
                
                # Check Data subfolder
                data_path = os.path.join(base_path, "Data")
                if os.path.exists(data_path):
                    is_valid, version, _, _ = self.validate_london_files(data_path)
                    if is_valid and version == "1.03":
                        logging.info(f"Found London 1.03 files at: {data_path}")
                        return base_path  # Return parent path
            
            logging.info("No London 1.03 files found on system")
            return None
            
        except Exception as e:
            logging.error(f"Error scanning for London 1.03 files: {e}")
            return None

    def start_update_process(self):
        """Start the update process for existing installation"""
        # Hide all widgets
        for widget in self.root.winfo_children():
            if not isinstance(widget, tk.Frame) or widget.winfo_class() != "Frame":
                continue
            for child in widget.winfo_children():
                child.pack_forget()
        
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Add custom title bar
        self.add_custom_title_bar()
        
        bg_color = "#1e1e1e"
        fg_color = "#ffffff"
        
        if self.bg_image:
            bg_label = tk.Label(self.root, image=self.bg_image)
            bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            bg_label.lower()
            main_container = tk.Frame(self.root)
            content_frame = tk.Frame(main_container)
        else:
            self.root.configure(bg=bg_color)
            main_container = tk.Frame(self.root, bg=bg_color)
            content_frame = tk.Frame(main_container, bg=bg_color)
        
        main_container.pack(fill="both", expand=True, padx=20, pady=5)
        content_frame.pack(fill="both", expand=True)
        
        if self.logo:
            tk.Label(content_frame, image=self.logo, bg=bg_color if not self.bg_image else '#1e1e1e').pack(pady=self.get_scaled_value(5))
        
        tk.Label(content_frame, text="Updating Fallout: London VR", font=self.title_font,
                bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color).pack(pady=self.get_scaled_value(5))
        
        self.message_label = tk.Label(content_frame, text="Preparing update...", font=self.regular_font,
                                      bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color,
                                      wraplength=460, justify="center")
        self.message_label.pack(pady=self.get_scaled_value(2))
        
        # Initialize and start slideshow
        self.initialize_and_start_slideshow()
        
        # Start update in background thread
        threading.Thread(target=self.perform_update, daemon=True).start()

    def handle_files_in_use(self, operation_name="operation"):
        """Show dialog when files are in use and wait for user to resolve
        
        Returns:
            bool: True if user wants to retry, False if user wants to cancel
        """
        message = (
            f"Files are currently in use and cannot be modified during the {operation_name}.\n\n"
            "Please:\n"
            "• Close Mod Organizer 2\n"
            "• Close Fallout 4 VR\n"
            "• Close any file explorers viewing the installation directory\n"
            "• Close any other programs that might be accessing the files\n\n"
            "Click 'Retry' when ready to continue, or 'Cancel' to abort the update."
        )
        
        # Use askretrycancel which returns True for Retry, False for Cancel
        result = messagebox.askretrycancel(
            "Files In Use - Action Required",
            message,
            icon='warning'
        )
        
        if result:
            logging.info(f"User chose to retry {operation_name} after closing files")
            return True
        else:
            logging.info(f"User cancelled update due to files in use during {operation_name}")
            return False

    def merge_modlist_txt(self, install_path, source_modlist_path):
        """Merge new mods from source modlist.txt into existing user modlist.txt
        
        This preserves user's existing mod order and any custom mods they've added,
        while inserting any new mods from the update in their relative positions.
        
        Args:
            install_path: Path to the existing Fallout London VR installation
            source_modlist_path: Path to the new modlist.txt from the update package
        """
        try:
            user_modlist_path = os.path.join(install_path, "profiles", "Default", "modlist.txt")
            
            if not os.path.exists(user_modlist_path):
                # No existing modlist, just copy the source
                logging.info("No existing modlist.txt found, copying source directly")
                os.makedirs(os.path.dirname(user_modlist_path), exist_ok=True)
                shutil.copy2(source_modlist_path, user_modlist_path)
                return
            
            if not os.path.exists(source_modlist_path):
                logging.warning(f"Source modlist.txt not found at {source_modlist_path}")
                return
            
            # Read both modlists
            with open(user_modlist_path, 'r', encoding='utf-8') as f:
                user_lines = f.readlines()
            
            with open(source_modlist_path, 'r', encoding='utf-8') as f:
                source_lines = f.readlines()
            
            # Parse mod entries (extract mod name without +/- prefix)
            def parse_mod_name(line):
                line = line.strip()
                if line.startswith('+') or line.startswith('-'):
                    return line[1:].strip()
                return None
            
            # Get list of mod names from user's file (preserving order)
            user_mods = []
            for line in user_lines:
                mod_name = parse_mod_name(line)
                if mod_name:
                    user_mods.append(mod_name)
            
            # Get ordered list of mod names from source file
            source_mods = []
            source_mod_lines = {}  # Map mod name to full line (with +/- prefix)
            for line in source_lines:
                mod_name = parse_mod_name(line)
                if mod_name:
                    source_mods.append(mod_name)
                    source_mod_lines[mod_name] = line.strip()
            
            # Find mods in source that are not in user's file
            user_mod_set = set(user_mods)
            missing_mods = [m for m in source_mods if m not in user_mod_set]
            
            if not missing_mods:
                logging.info("No new mods to add to modlist.txt")
                return
            
            logging.info(f"Adding {len(missing_mods)} new mods to modlist.txt: {missing_mods}")
            
            # Build new modlist by inserting missing mods in their relative positions
            result_lines = []
            user_line_index = 0
            
            # Copy header lines (comments at the start)
            for line in user_lines:
                if line.strip().startswith('#') or not line.strip():
                    result_lines.append(line)
                else:
                    break
            
            # Skip header lines we already added
            header_count = len(result_lines)
            
            # Process remaining lines
            for i, line in enumerate(user_lines[header_count:], start=header_count):
                mod_name = parse_mod_name(line)
                
                # Before adding this user mod, check if any missing mods should go before it
                if mod_name:
                    for missing_mod in missing_mods[:]:  # Copy list to allow modification
                        # Find where this missing mod appears in source relative to current user mod
                        try:
                            missing_idx = source_mods.index(missing_mod)
                            current_idx = source_mods.index(mod_name) if mod_name in source_mods else len(source_mods)
                            
                            if missing_idx < current_idx:
                                # This missing mod should appear before current mod
                                result_lines.append(source_mod_lines[missing_mod] + '\n')
                                missing_mods.remove(missing_mod)
                                logging.info(f"Inserted mod '{missing_mod}' before '{mod_name}'")
                        except ValueError:
                            continue
                
                result_lines.append(line)
            
            # Add any remaining missing mods at the end (before Fallout London Data and Fallout 4 Data if possible)
            for missing_mod in missing_mods:
                result_lines.append(source_mod_lines[missing_mod] + '\n')
                logging.info(f"Appended mod '{missing_mod}' at end")
            
            # Write merged modlist
            with open(user_modlist_path, 'w', encoding='utf-8') as f:
                f.writelines(result_lines)
            
            logging.info(f"Successfully merged modlist.txt")
            
        except Exception as e:
            logging.error(f"Failed to merge modlist.txt: {e}")
            # Non-fatal, continue with update

    def read_f4vr_path_from_mo2_ini(self, install_path):
        """Read the Fallout 4 VR path from ModOrganizer.ini
        
        Args:
            install_path: Path to the existing Fallout London VR installation
            
        Returns:
            str: The F4VR path if found, None otherwise
        """
        try:
            mo2_ini_path = os.path.join(install_path, "ModOrganizer.ini")
            if not os.path.exists(mo2_ini_path):
                logging.warning(f"ModOrganizer.ini not found at {mo2_ini_path}")
                return None
            
            config = configparser.ConfigParser()
            config.read(mo2_ini_path, encoding='utf-8')
            
            if 'General' in config and 'gamePath' in config['General']:
                # MO2 stores paths with forward slashes, convert to backslashes for Windows
                f4vr_path = config['General']['gamePath'].replace('/', '\\')
                
                # Strip @ByteArray(...) wrapper if present (MO2/Qt encoding)
                if f4vr_path.startswith('@ByteArray(') and f4vr_path.endswith(')'):
                    f4vr_path = f4vr_path[11:-1]  # Remove '@ByteArray(' prefix and ')' suffix
                    logging.info(f"Stripped @ByteArray wrapper from path: {f4vr_path}")
                
                # Verify the path exists and contains Fallout4VR.exe
                if os.path.exists(f4vr_path):
                    f4vr_exe = os.path.join(f4vr_path, "Fallout4VR.exe")
                    if os.path.exists(f4vr_exe):
                        logging.info(f"Read F4VR path from ModOrganizer.ini: {f4vr_path}")
                        return f4vr_path
                    else:
                        logging.warning(f"Fallout4VR.exe not found at {f4vr_path}")
                else:
                    logging.warning(f"F4VR path from ModOrganizer.ini does not exist: {f4vr_path}")
            else:
                logging.warning("gamePath not found in ModOrganizer.ini [General] section")
            
            return None
            
        except Exception as e:
            logging.error(f"Error reading F4VR path from ModOrganizer.ini: {e}")
            return None

    def perform_update(self):
        """Perform update of existing installation"""
        try:
            install_path = self.mo2_path.get()
            
            # Read F4VR path from ModOrganizer.ini for CAS and xSE Preloader installation
            f4vr_path = self.read_f4vr_path_from_mo2_ini(install_path)
            if f4vr_path:
                self.f4vr_path.set(f4vr_path)
                logging.info(f"Set f4vr_path from ModOrganizer.ini: {f4vr_path}")
            else:
                logging.warning("Could not read F4VR path from ModOrganizer.ini - CAS and xSE Preloader may be skipped")
            
            self.root.after(0, lambda: self.message_label.config(text="Backing up .ini files...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Step 1: Backup .ini files from Default profile
            profiles_dir = os.path.join(install_path, "profiles")
            default_profile_path = os.path.join(profiles_dir, "Default")
            if os.path.exists(default_profile_path):
                ini_files = ["fallout4.ini", "fallout4prefs.ini", "fallout4custom.ini"]
                for ini_file in ini_files:
                    ini_path = os.path.join(default_profile_path, ini_file)
                    if os.path.exists(ini_path):
                        # Create backup name: fallout4.ini -> fallout4old.ini
                        backup_name = ini_file.replace("fallout4", "fallout4old")
                        backup_ini_path = os.path.join(default_profile_path, backup_name)
                        shutil.copy2(ini_path, backup_ini_path)
                        logging.info(f"Backed up {ini_file} to {backup_name}")
            
            # Step 1.5: Remove deprecated files and mods BEFORE copying new assets
            self.root.after(0, lambda: self.message_label.config(text="Removing deprecated mods...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            mods_dir = os.path.join(install_path, "mods")
            logging.info(f"Checking mods directory for cleanup: {mods_dir}")
            
            # List what's in mods directory before cleanup
            if os.path.exists(mods_dir):
                mods_list = os.listdir(mods_dir)
                logging.info(f"Mods found before cleanup: {mods_list}")
            else:
                logging.warning(f"Mods directory does not exist: {mods_dir}")
            
            # Check for old mod organization (pre-0.96) - if Plugins folder contains DLLs, remove entire F4SE folder
            f4se_plugins_path = os.path.join(install_path, "mods", "Fallout London VR", "F4SE", "Plugins")
            f4se_folder_path = os.path.join(install_path, "mods", "Fallout London VR", "F4SE")
            if os.path.exists(f4se_plugins_path):
                try:
                    # Check if any DLL files exist in the Plugins folder
                    dll_files = [f for f in os.listdir(f4se_plugins_path) if f.lower().endswith('.dll')]
                    if dll_files:
                        logging.info(f"Detected old mod organization (pre-0.96): Found DLLs in Plugins folder: {dll_files}")
                        self.root.after(0, lambda: self.message_label.config(text="Cleaning up old mod organization...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        
                        while True:  # Retry loop
                            try:
                                shutil.rmtree(f4se_folder_path)
                                logging.info(f"Removed old F4SE folder from Fallout London VR mod: {f4se_folder_path}")
                                break  # Success, exit retry loop
                            except PermissionError as e:
                                logging.warning(f"Permission error removing old F4SE folder: {e}")
                                if not self.handle_files_in_use("folder removal"):
                                    self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                                    return
                                # User clicked retry, loop continues
                            except Exception as e:
                                logging.error(f"Failed to remove old F4SE folder: {e}")
                                break  # Non-permission error, continue update
                except Exception as e:
                    logging.warning(f"Error checking for old mod organization: {e}")
            
            # Remove any old FRIK directories and install new FRIK
            self.root.after(0, lambda: self.message_label.config(text="Updating FRIK...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Find and remove any existing FRIK directories (FRIK.v74, FRIK.v75, FRIK.v76, etc.)
            frik_dirs_removed = False
            try:
                for item in os.listdir(mods_dir):
                    item_path = os.path.join(mods_dir, item)
                    # Match any directory starting with "FRIK" (case-insensitive)
                    if os.path.isdir(item_path) and item.upper().startswith("FRIK"):
                        while True:  # Retry loop for directory removal
                            try:
                                self.root.after(0, lambda name=item: self.message_label.config(text=f"Removing old FRIK: {name}...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                                logging.info(f"Found old FRIK directory at {item_path}, removing...")
                                shutil.rmtree(item_path)
                                logging.info(f"Old FRIK directory '{item}' removed successfully")
                                frik_dirs_removed = True
                                break  # Success, exit retry loop
                            except PermissionError as e:
                                logging.warning(f"Permission error removing {item}: {e}")
                                # Show retry/cancel dialog
                                if not self.handle_files_in_use("directory removal"):
                                    # User cancelled
                                    self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                                    return
                                # User clicked retry, loop continues
                            except Exception as e:
                                logging.error(f"Failed to remove old FRIK directory '{item}': {e}")
                                self.root.after(0, lambda es=str(e): messagebox.showerror("Error", f"Failed to remove old FRIK version: {es}"))
                                break  # Continue with other directories
            except Exception as e:
                logging.warning(f"Error scanning for FRIK directories: {e}")
            
            # Always download and install new FRIK
            self.root.after(0, lambda: self.message_label.config(text="Installing latest FRIK version...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            try:
                self.download_and_install_frik()
                logging.info("New FRIK version installed successfully")
            except Exception as e:
                logging.warning(f"Failed to install FRIK: {e}")
                # Non-fatal, continue with update
            
            # Check and remove High FPS Physics Fix if present
            high_fps_path = os.path.join(mods_dir, "High FPS Physics Fix")
            if os.path.exists(high_fps_path):
                while True:  # Retry loop
                    try:
                        self.root.after(0, lambda: self.message_label.config(text="Removing deprecated High FPS Physics Fix...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        logging.info(f"Found High FPS Physics Fix at {high_fps_path}, removing...")
                        shutil.rmtree(high_fps_path)
                        logging.info("High FPS Physics Fix removed successfully")
                        break  # Success, exit retry loop
                    except PermissionError as e:
                        logging.warning(f"Permission error removing High FPS Physics Fix: {e}")
                        # Show retry/cancel dialog
                        if not self.handle_files_in_use("directory removal"):
                            # User cancelled
                            self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                            return
                        # User clicked retry, loop continues
                    except Exception as e:
                        logging.warning(f"Failed to remove High FPS Physics Fix: {e}")
                        # Non-fatal, continue with update
                        break
            
            # Check and remove XDI mod directory if present
            xdi_mod_path = os.path.join(mods_dir, "XDI")
            if os.path.exists(xdi_mod_path):
                while True:  # Retry loop
                    try:
                        self.root.after(0, lambda: self.message_label.config(text="Removing deprecated XDI mod...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        logging.info(f"Found XDI mod directory at {xdi_mod_path}, removing...")
                        shutil.rmtree(xdi_mod_path)
                        logging.info("XDI mod directory removed successfully")
                        break  # Success, exit retry loop
                    except PermissionError as e:
                        logging.warning(f"Permission error removing XDI mod: {e}")
                        # Show retry/cancel dialog
                        if not self.handle_files_in_use("directory removal"):
                            # User cancelled
                            self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                            return
                        # User clicked retry, loop continues
                    except Exception as e:
                        logging.warning(f"Failed to remove XDI mod directory: {e}")
                        # Non-fatal, continue with update
                        break
            
            # Check and remove PrivateProfileRedirector F4 if present
            ppr_mod_path = os.path.join(mods_dir, "PrivateProfileRedirector F4")
            if os.path.exists(ppr_mod_path):
                while True:  # Retry loop
                    try:
                        self.root.after(0, lambda: self.message_label.config(text="Removing deprecated PrivateProfileRedirector...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        logging.info(f"Found PrivateProfileRedirector F4 at {ppr_mod_path}, removing...")
                        shutil.rmtree(ppr_mod_path)
                        logging.info("PrivateProfileRedirector F4 removed successfully")
                        break  # Success, exit retry loop
                    except PermissionError as e:
                        logging.warning(f"Permission error removing PrivateProfileRedirector F4: {e}")
                        # Show retry/cancel dialog
                        if not self.handle_files_in_use("directory removal"):
                            # User cancelled
                            self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                            return
                        # User clicked retry, loop continues
                    except Exception as e:
                        logging.warning(f"Failed to remove PrivateProfileRedirector F4: {e}")
                        # Non-fatal, continue with update
                        break
            
            # Check and remove Version Check Patcher if present (check both possible folder names)
            for vcheck_folder in ["Version Check Patcher", "vcheck_patcher"]:
                version_check_patcher_path = os.path.join(mods_dir, vcheck_folder)
                if os.path.exists(version_check_patcher_path):
                    while True:  # Retry loop
                        try:
                            self.root.after(0, lambda: self.message_label.config(text="Removing deprecated Version Check Patcher...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                            logging.info(f"Found Version Check Patcher at {version_check_patcher_path}, removing...")
                            shutil.rmtree(version_check_patcher_path)
                            logging.info("Version Check Patcher removed successfully")
                            break  # Success, exit retry loop
                        except PermissionError as e:
                            logging.warning(f"Permission error removing Version Check Patcher: {e}")
                            # Show retry/cancel dialog
                            if not self.handle_files_in_use("directory removal"):
                                # User cancelled
                                self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                                return
                            # User clicked retry, loop continues
                        except Exception as e:
                            logging.warning(f"Failed to remove Version Check Patcher: {e}")
                            # Non-fatal, continue with update
                            break
            
            # Step 1.9: Remove old Fallout London VR mod folder before copying new assets
            fallout_london_vr_mod_path = os.path.join(mods_dir, "Fallout London VR")
            if os.path.exists(fallout_london_vr_mod_path):
                while True:  # Retry loop
                    try:
                        self.root.after(0, lambda: self.message_label.config(text="Removing old Fallout London VR mod...", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        logging.info(f"Found Fallout London VR mod at {fallout_london_vr_mod_path}, removing before update...")
                        shutil.rmtree(fallout_london_vr_mod_path)
                        logging.info("Fallout London VR mod folder removed successfully")
                        break  # Success, exit retry loop
                    except PermissionError as e:
                        logging.warning(f"Permission error removing Fallout London VR mod: {e}")
                        # Show retry/cancel dialog
                        if not self.handle_files_in_use("folder removal"):
                            # User cancelled
                            self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                            return
                        # User clicked retry, loop continues
                    except Exception as e:
                        logging.error(f"Failed to remove Fallout London VR mod folder: {e}")
                        self.root.after(0, lambda es=str(e): messagebox.showerror("Error", f"Failed to remove old Fallout London VR mod: {es}"))
                        return  # Fatal error, can't continue without removing old folder
            
            # Step 2: Copy MO2 assets (update mods)
            self.root.after(0, lambda: self.message_label.config(text="Copying updated mod files...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            while True:  # Retry loop for copying assets
                try:
                    # Exclude ModOrganizer.ini and modlist.txt - these need special handling
                    self.copy_mo2_assets(update_config=False, exclude_files=["ModOrganizer.ini", "modlist.txt"])
                    break  # Success, exit retry loop
                except PermissionError as e:
                    logging.warning(f"Permission error copying MO2 assets: {e}")
                    # Show retry/cancel dialog
                    if not self.handle_files_in_use("file copying"):
                        # User cancelled
                        self.root.after(0, lambda: self.message_label.config(text="Update cancelled by user", fg="#ff6666") if self.message_label.winfo_exists() else None)
                        return
                    # User clicked retry, loop continues
                except Exception as e:
                    error_msg = str(e)
                    logging.error(f"Failed to copy MO2 assets: {error_msg}")
                    self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Failed to copy mod files: {msg}"))
                    return
            
            # Step 2.1: Merge modlist.txt (add any new mods while preserving user's order)
            self.root.after(0, lambda: self.message_label.config(text="Updating mod list...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            try:
                # Extract source modlist.txt from MO2.7z to temp location for merge
                mo2_assets_archive = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "MO2.7z")
                if os.path.exists(mo2_assets_archive):
                    temp_dir = tempfile.mkdtemp(prefix="folvr_modlist_")
                    bundled_7za = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "7za.exe")
                    
                    # Extract just the modlist.txt file
                    extract_cmd = [bundled_7za, "e", mo2_assets_archive, f"-o{temp_dir}", "MO2/profiles/Default/modlist.txt", "-y", "-bb0", "-bd"]
                    result = subprocess.run(extract_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    source_modlist = os.path.join(temp_dir, "modlist.txt")
                    if os.path.exists(source_modlist):
                        self.merge_modlist_txt(install_path, source_modlist)
                    else:
                        logging.warning("Could not extract modlist.txt from MO2.7z for merge")
                    
                    # Clean up temp dir
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass
            except Exception as e:
                logging.warning(f"Failed to merge modlist.txt: {e}")
                # Non-fatal, continue with update
            
            # Step 2.5: Upgrade to London 1.03 if user opted in
            if self.upgrade_to_103 and self.london_103_source_path:
                self.root.after(0, lambda: self.message_label.config(text="Upgrading to Fallout: London 1.03", fg="#ffffff") if self.message_label.winfo_exists() else None)
                try:
                    # Destination is mods/Fallout London Data
                    london_data_dest = os.path.join(install_path, "mods", "Fallout London Data")
                    
                    # Determine source data directory
                    if os.path.exists(os.path.join(self.london_103_source_path, "Data", "LondonWorldSpace.esm")):
                        source_data_dir = os.path.join(self.london_103_source_path, "Data")
                    elif os.path.exists(os.path.join(self.london_103_source_path, "LondonWorldSpace.esm")):
                        source_data_dir = self.london_103_source_path
                    else:
                        raise FileNotFoundError("Could not find London files in source path")
                    
                    logging.info(f"Copying London 1.03 files from {source_data_dir} to {london_data_dest}")
                    
                    # Copy London files using existing method
                    self.copy_london_files_only(os.path.dirname(source_data_dir) if source_data_dir.endswith("Data") else source_data_dir, london_data_dest)
                    
                    logging.info("London 1.03 upgrade completed successfully")
                except Exception as e:
                    logging.error(f"Failed to upgrade to London 1.03: {e}")
                    self.root.after(0, lambda es=str(e): messagebox.showwarning("Warning", f"Failed to upgrade to London 1.03: {es}\n\nContinuing with update..."))
                    # Non-fatal, continue with update
            
            # Step 3: Update fallout4custom.ini with VRUI settings
            try:
                default_profile_path = os.path.join(install_path, "profiles", "Default")
                custom_ini_path = os.path.join(default_profile_path, "fallout4custom.ini")
                
                # Read or create the custom INI file
                config = configparser.ConfigParser(strict=False)
                if os.path.exists(custom_ini_path):
                    config.read(custom_ini_path, encoding='utf-8')
                    logging.info(f"Read existing fallout4custom.ini")
                else:
                    logging.info(f"Creating new fallout4custom.ini at {custom_ini_path}")
                
                # Add or update [VRUI] section
                if 'VRUI' not in config:
                    config['VRUI'] = {}
                
                config['VRUI']['iVRUIRenderTargetHeight'] = '4096'
                config['VRUI']['iVRUIRenderTargetWidth'] = '4096'
                
                # Write back to file
                with open(custom_ini_path, 'w', encoding='utf-8') as configfile:
                    config.write(configfile)
                
                logging.info(f"Updated fallout4custom.ini with VRUI settings")
            except Exception as e:
                logging.warning(f"Failed to update fallout4custom.ini: {e}")
                # Non-fatal, continue with update
            
            # Step 4: Update FRIK weapon offsets
            self.root.after(0, lambda: self.message_label.config(text="Updating FRIK weapon offsets...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            try:
                self.copy_weapon_offsets()
                logging.info("FRIK weapon offsets updated successfully")
            except Exception as e:
                logging.warning(f"Failed to update FRIK weapon offsets: {e}")
                # Non-fatal, continue with update
            
            # Step 5: Install xSE Plugin Preloader
            try:
                self.xse_preloader_installed = self.install_xse_plugin_preloader()
            except Exception as e:
                logging.warning(f"Failed to install xSE Plugin Preloader: {e}")
                self.xse_preloader_installed = False
                # Non-fatal, continue with update
            
            # Step 7: Complete
            self.root.after(0, lambda: self.message_label.config(text="Update completed successfully!", fg="#00ff00") if self.message_label.winfo_exists() else None)
            logging.info("Update completed successfully")
            
            # Show completion UI
            self.show_completion_ui()
            
        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Update failed: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            logging.error(f"Update failed: {e}")
            raise

    def create_welcome_page(self):
        """Create the welcome page with custom title bar"""
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Add custom title bar first
        self.add_custom_title_bar()

        bg_color = "#1e1e1e"
        fg_color = "#ffffff"
        entry_fg_color = "#000000"
        entry_bg_color = "#d3d3d3"
        accent_color = "#0078d7"
        if self.bg_image:
            bg_label = tk.Label(self.root, image=self.bg_image)
            bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            bg_label.lower()
            # Use default background (no bg set) for transparency effect
            main_container = tk.Frame(self.root)
        else:
            self.root.configure(bg=bg_color)
            main_container = tk.Frame(self.root, bg=bg_color)
        main_container.pack(fill="both", expand=True, padx=self.get_scaled_value(20), pady=(self.get_scaled_value(5), 0))
        
        # Create a canvas with scrollbar for small screens
        canvas = tk.Canvas(main_container, bg=bg_color if not self.bg_image else '#1e1e1e', 
                          highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        # Create scrollable frame inside canvas
        scrollable_frame = tk.Frame(canvas, bg=bg_color if not self.bg_image else '#1e1e1e')
        
        # Configure canvas scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="n")
        
        # Make the scrollable frame expand to canvas width
        def configure_scroll_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
            # Update scroll region after resize
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.bind('<Configure>', configure_scroll_width)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas only (scrollbar hidden but still functional for mousewheel)
        canvas.pack(side="left", fill="both", expand=True)
        # scrollbar.pack(side="right", fill="y")  # Hidden but scrollbar still exists for functionality
        
        # Schedule scroll region update after all widgets are packed
        def update_scroll_region():
            canvas.configure(scrollregion=canvas.bbox("all"))
        self.root.after(100, update_scroll_region)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Enable arrow key scrolling
        def on_key_up(event):
            canvas.yview_scroll(-1, "units")
        def on_key_down(event):
            canvas.yview_scroll(1, "units")
        canvas.bind_all("<Up>", on_key_up)
        canvas.bind_all("<Down>", on_key_down)
        
        # Store canvas reference for cleanup
        self.welcome_canvas = canvas
        
        # Header section - logo and title (in scrollable frame)
        header_frame = tk.Frame(scrollable_frame, bg=bg_color if not self.bg_image else '#1e1e1e')
        header_frame.pack(fill="x", pady=(0, 0))
        
        if self.logo:
            tk.Label(header_frame, image=self.logo, bg=bg_color if not self.bg_image else '#1e1e1e').pack(pady=self.get_scaled_value(5))
        else:
            tk.Label(header_frame, text="[Logo Not Found]", font=self.bold_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color).pack(pady=self.get_scaled_value(5))
        self.header_label = tk.Label(header_frame, text="Fallout: London VR Installation", font=self.title_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color)
        self.header_label.pack(pady=self.get_scaled_value(5))
        
        # Content section - now inside scrollable frame
        content_frame = tk.Frame(scrollable_frame, bg=bg_color if not self.bg_image else '#1e1e1e')
        content_frame.pack(fill="both", expand=True)
        self.content_frame = content_frame
        
        self.dlc_status_label = tk.Label(content_frame, text="", font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color, wraplength=self.get_scaled_value(460), justify="center", anchor="center")
        self.dlc_status_label.pack(pady=self.get_scaled_value(2), fill="x")
        
        self.f4vr_status_label = tk.Label(content_frame, text="", font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color, wraplength=self.get_scaled_value(460), justify="center", anchor="center")
        self.f4vr_status_label.pack(pady=self.get_scaled_value(2), fill="x")
        
        self.london_status_label = tk.Label(content_frame, text="", font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color, wraplength=self.get_scaled_value(460), justify="center", anchor="center")
        self.london_status_label.pack(pady=self.get_scaled_value(2), fill="x")
        self.disk_space_label = tk.Label(content_frame, text="", font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color, anchor="center")
        self.disk_space_label.pack(pady=self.get_scaled_value(2), fill="x")
        self.message_label = tk.Label(content_frame, text="", font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color, wraplength=self.get_scaled_value(460), justify="center", anchor="center")
        self.message_label.pack(pady=self.get_scaled_value(2), fill="x")
        
        # London path section (shown FIRST in fresh install)
        self.london_label = tk.Label(content_frame, text="Fallout London Location", font=self.bold_font,
                                     bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color)
        self.london_label.pack(pady=(self.get_scaled_value(8), self.get_scaled_value(3)))
        self.london_data_path.set("")
        self.london_data_entry = tk.Entry(content_frame, textvariable=self.london_data_path, width=40,
                                          font=self.regular_font, bg=entry_bg_color, fg=entry_fg_color,
                                          insertbackground=fg_color, bd=0)
        self.london_data_entry.pack(pady=self.get_scaled_value(5))
        self.london_browse_button = tk.Button(content_frame, text="Browse",
                                              command=lambda: self.browse_path(self.london_data_path),
                                              font=self.regular_font, bg=accent_color, fg=fg_color, bd=0,
                                              relief="flat", activebackground="#005ba1", padx=10, pady=5)
        self.london_browse_button.pack()
        
        # Fallout 4 VR path section (shown SECOND)
        self.f4vr_label = tk.Label(content_frame, text="Fallout 4 VR Location", font=self.bold_font,
                 bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color)
        self.f4vr_label.pack(pady=(self.get_scaled_value(8), self.get_scaled_value(3)))
        self.f4vr_entry = tk.Entry(content_frame, textvariable=self.f4vr_path, width=40, font=self.regular_font,
                                   bg=entry_bg_color, fg=entry_fg_color, insertbackground=fg_color, bd=0)
        self.f4vr_entry.pack(pady=self.get_scaled_value(5))
        self.f4vr_browse_button = tk.Button(content_frame, text="Browse", command=lambda: self.browse_path(self.f4vr_path),
                                            font=self.regular_font, bg=accent_color, fg=fg_color, bd=0,
                                            relief="flat", activebackground="#005ba1", padx=self.get_scaled_value(10), pady=self.get_scaled_value(5))
        self.f4vr_browse_button.pack()
        
        # Fallout 4 DLC path section (hidden initially, shown only if DLC not found in London or F4VR)
        self.f4_label = tk.Label(content_frame, text="Fallout 4 DLC Location", font=self.bold_font,
                 bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color)
        self.f4_entry = tk.Entry(content_frame, textvariable=self.f4_path, width=40, font=self.regular_font,
                                 bg=entry_bg_color, fg=entry_fg_color, insertbackground=fg_color, bd=0)
        self.f4_browse_button = tk.Button(content_frame, text="Browse", command=lambda: self.browse_path(self.f4_path),
                                          font=self.regular_font, bg=accent_color, fg=fg_color, bd=0,
                                          relief="flat", activebackground="#005ba1", padx=self.get_scaled_value(10), pady=self.get_scaled_value(5))
        # F4 DLC widgets are NOT packed initially - they will be shown by show_f4_dlc_widgets() if needed
        
        # Atkins image label (shown for update mode, hidden for fresh install)
        self.atkins_label = tk.Label(content_frame, image=self.atkins_image if self.atkins_image else None, 
                                     bg=bg_color if not self.bg_image else '#1e1e1e')
        # Don't pack yet, will be packed in _reorganize_ui_for_update
        
        # Create a separate frame for installation directory
        self.installation_dir_container = tk.Frame(content_frame, bg=bg_color if not self.bg_image else '#1e1e1e')
        self.installation_dir_label = tk.Label(self.installation_dir_container, text="Installation Location", font=self.bold_font,
                                               bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color)
        self.installation_dir_label.pack(pady=(self.get_scaled_value(8), self.get_scaled_value(3)))
        # Use initial_install_path if provided (when restarting in fresh install mode), otherwise use default
        if self.initial_install_path:
            self.mo2_path.set(self.initial_install_path)
        else:
            default_path = os.path.join("C:\\", "Games", "Fallout London VR")
            self.mo2_path.set(default_path)
        self.installation_entry = tk.Entry(self.installation_dir_container, textvariable=self.mo2_path, width=40,
                                           font=self.regular_font, bg=entry_bg_color, fg=entry_fg_color,
                                           insertbackground=fg_color, bd=0)
        self.installation_entry.pack(pady=self.get_scaled_value(3))
        self.installation_browse_button = tk.Button(self.installation_dir_container, text="Browse",
                                                    command=self.browse_mo2_path, font=self.regular_font,
                                                    bg=accent_color, fg=fg_color, bd=0, relief="flat",
                                                    activebackground="#005ba1", padx=self.get_scaled_value(10), pady=self.get_scaled_value(3))
        self.installation_browse_button.pack(pady=self.get_scaled_value(3))
        # Pack in content_frame for now (fresh install mode)
        self.installation_dir_container.pack(pady=(0, 2))
        
        self.update_disk_space() # Immediately update disk space display
        
        # Add status message for existing installation detection
        self.installation_status_label = tk.Label(content_frame, text="", font=self.regular_font, 
                                                  bg=bg_color if not self.bg_image else '#1e1e1e', fg="#ffa500",
                                                  wraplength=self.get_scaled_value(460), justify="center", anchor="center")
        self.installation_status_label.pack(pady=self.get_scaled_value(2), fill="x")
        
        # Instruction label for fresh install mode - created in content_frame like update mode
        self.instruction_label = tk.Label(content_frame, text="Press Install or Enter to continue", 
                                         font=self.regular_font, bg=bg_color if not self.bg_image else '#1e1e1e', 
                                         fg="#ffffff")
        self.instruction_label.pack(pady=(self.get_scaled_value(10), self.get_scaled_value(3)))
        
        self.install_button = tk.Button(content_frame, text="Install", command=self.validate_and_install,
                                        font=self.regular_font, bg="#28a745", fg=fg_color, bd=0,
                                        relief="flat", activebackground="#218838", padx=self.get_scaled_value(10), pady=self.get_scaled_value(5))
        self.install_button.pack(pady=(self.get_scaled_value(2), self.get_scaled_value(10)))
        
        # Bind Enter key to trigger Install button
        self.root.bind('<Return>', lambda event: self.validate_and_install())
        
        self.root.after(100, self.detect_paths)
        # Set trace for MO2 path to trigger disk space check AND detect manual changes to switch back to fresh install
        self.mo2_trace_id = self.mo2_path.trace_add("write", lambda *args: self.on_installation_path_changed())

    def on_f4_path_change(self):
        """Handle Fallout 4 path changes"""
        f4_path = self.f4_path.get()
        if f4_path and os.path.exists(f4_path):
            f4_exe = os.path.join(f4_path, "Fallout4.exe")
            if os.path.exists(f4_exe):
                self.check_london_installation(f4_path)
                self.update_disk_space()

    def hide_london_widgets(self):
        """Hide London data widgets"""
        try:
            if hasattr(self, 'london_label'):
                self.london_label.pack_forget()
            if hasattr(self, 'london_data_entry'):
                self.london_data_entry.pack_forget()
            if hasattr(self, 'london_browse_button'):
                self.london_browse_button.pack_forget()
        except Exception as e:
            logging.warning(f"Error hiding London widgets: {e}")

    def hide_f4_dlc_widgets(self):
        """Hide Fallout 4 DLC path widgets"""
        try:
            if hasattr(self, 'f4_label'):
                self.f4_label.pack_forget()
            if hasattr(self, 'f4_entry'):
                self.f4_entry.pack_forget()
            if hasattr(self, 'f4_browse_button'):
                self.f4_browse_button.pack_forget()
        except Exception as e:
            logging.warning(f"Error hiding F4 DLC widgets: {e}")

    def validate_london_files(self, path):
        """
        Validate Fallout: London installation files and detect version.
        
        Returns tuple: (is_valid, version, status_message, missing_files)
        - is_valid: True if all required files are present
        - version: "1.03" or "1.02" or None
        - status_message: Message to display
        - missing_files: List of missing files (empty if valid)
        """
        # Required files for both versions (29 base files)
        base_required_files = [
            "LondonWorldSpace - Animations.ba2",
            "LondonWorldSpace - Interface.ba2",
            "LondonWorldSpace - Materials.ba2",
            "LondonWorldSpace - Meshes.ba2",
            "LondonWorldSpace - MeshesExtra.ba2",
            "LondonWorldSpace - MeshesLOD.ba2",
            "LondonWorldSpace - Misc.ba2",
            "LondonWorldSpace - Sounds.ba2",
            "LondonWorldSpace - Textures1.ba2",
            "LondonWorldSpace - Textures2.ba2",
            "LondonWorldSpace - Textures3.ba2",
            "LondonWorldSpace - Textures4.ba2",
            "LondonWorldSpace - Textures5.ba2",
            "LondonWorldSpace - Textures6.ba2",
            "LondonWorldSpace - Textures7.ba2",
            "LondonWorldSpace - Textures8.ba2",
            "LondonWorldSpace - Textures9.ba2",
            "LondonWorldSpace - Textures10.ba2",
            "LondonWorldSpace - Textures11.ba2",
            "LondonWorldSpace - Textures12.ba2",
            "LondonWorldSpace - Textures13.ba2",
            "LondonWorldSpace - Voices.ba2",
            "LondonWorldSpace - VoicesExtra.ba2",
            "LondonWorldSpace.cdx",
            "LondonWorldSpace.esm",
            "LondonWorldSpace-DLCBlock.esp",
        ]
        
        # Additional file for version 1.03
        v103_extra_file = "LondonWorldSpace - Textures14.ba2"
        
        # Required folders
        required_folders = ["Scripts", "Video"]
        
        # Optional files (not required for validation)
        optional_files = [
            "LondonWorldSpace - CleanUp.bat",
            "LondonWorldSpace - Credits.txt",
            "LondonWorldSpace - Geometry.csg",
        ]
        
        try:
            # First, find the data directory (could be root or Data subfolder)
            data_dir = None
            if os.path.exists(os.path.join(path, "LondonWorldSpace.esm")):
                data_dir = path
            elif os.path.exists(os.path.join(path, "Data", "LondonWorldSpace.esm")):
                data_dir = os.path.join(path, "Data")
            else:
                return (False, None, "Fallout: London: Invalid location", ["LondonWorldSpace.esm not found"])
            
            # Get list of files in the data directory (case-insensitive)
            try:
                dir_contents = os.listdir(data_dir)
                files_lower = {f.lower(): f for f in dir_contents if os.path.isfile(os.path.join(data_dir, f))}
                folders_lower = {f.lower(): f for f in dir_contents if os.path.isdir(os.path.join(data_dir, f))}
            except Exception as e:
                logging.error(f"Error listing directory {data_dir}: {e}")
                return (False, None, f"Error accessing Fallout: London location: {str(e)}", [])
            
            # Check for required base files
            missing_files = []
            for req_file in base_required_files:
                if req_file.lower() not in files_lower:
                    missing_files.append(req_file)
            
            # Check for required folders
            missing_folders = []
            for req_folder in required_folders:
                if req_folder.lower() not in folders_lower:
                    missing_folders.append(req_folder)
            
            # Check for version 1.03 extra file
            has_textures14 = v103_extra_file.lower() in files_lower
            
            # Determine result
            if missing_files or missing_folders:
                # Files are incomplete
                all_missing = missing_files + [f"{f}/" for f in missing_folders]
                logging.warning(f"Fallout: London files incomplete. Missing: {all_missing}")
                return (False, None, "Fallout: London files incomplete. Please redownload.", all_missing)
            
            # All required files present - determine version
            if has_textures14:
                version = "1.03"
                status_msg = "Fallout: London 1.03 Rabbit & Pork: Ready for installation"
            else:
                version = "1.02"
                status_msg = "Fallout: London 1.02: Ready for installation"
            
            logging.info(f"Fallout: London version {version} detected at {data_dir}")
            self.london_source_path = data_dir
            self.london_version = version
            
            return (True, version, status_msg, [])
            
        except Exception as e:
            logging.error(f"Error validating Fallout: London files: {e}")
            return (False, None, f"Error validating Fallout: London files: {str(e)}", [])

    def validate_f4vr_files(self, path):
        """
        Validate Fallout 4 VR installation files.
        Returns tuple: (is_valid, status_message, missing_files)
        
        Required files (28 total):
        - 21 Fallout4 base .ba2 files
        - 3 Fallout4_VR .ba2 files  
        - 2 .esm files
        - 1 .cdx file
        - 1 .csg file
        """
        required_files = [
            # Fallout4 base archives
            "Fallout4 - Animations.ba2",
            "Fallout4 - Interface.ba2",
            "Fallout4 - Materials.ba2",
            "Fallout4 - Meshes.ba2",
            "Fallout4 - MeshesExtra.ba2",
            "Fallout4 - Misc - Beta.ba2",
            "Fallout4 - Misc - Debug.ba2",
            "Fallout4 - Misc.ba2",
            "Fallout4 - Shaders.ba2",
            "Fallout4 - Sounds.ba2",
            "Fallout4 - Startup.ba2",
            "Fallout4 - Textures1.ba2",
            "Fallout4 - Textures2.ba2",
            "Fallout4 - Textures3.ba2",
            "Fallout4 - Textures4.ba2",
            "Fallout4 - Textures5.ba2",
            "Fallout4 - Textures6.ba2",
            "Fallout4 - Textures7.ba2",
            "Fallout4 - Textures8.ba2",
            "Fallout4 - Textures9.ba2",
            "Fallout4 - Voices.ba2",
            # VR-specific archives
            "Fallout4_VR - Main.ba2",
            "Fallout4_VR - Shaders.ba2",
            "Fallout4_VR - Textures.ba2",
            # ESM files
            "Fallout4.esm",
            "Fallout4_VR.esm",
            # Other required files
            "Fallout4.cdx",
            "Fallout4 - Geometry.csg",
        ]
        
        try:
            # Check if path exists
            if not os.path.exists(path):
                return (False, "Fallout 4 VR: Invalid location", [])
            
            # Determine where data files should be (root or Data folder)
            data_dir = path
            data_folder = os.path.join(path, "Data")
            
            # Check if Data folder exists and contains the files
            if os.path.exists(data_folder):
                # Check if files are in Data folder
                test_file = os.path.join(data_folder, "Fallout4.esm")
                if os.path.exists(test_file):
                    data_dir = data_folder
            
            # Get list of files in data directory (case-insensitive matching)
            try:
                existing_files = os.listdir(data_dir)
                existing_files_lower = {f.lower(): f for f in existing_files}
            except Exception as e:
                logging.error(f"Error listing F4VR data directory {data_dir}: {e}")
                return (False, f"Error accessing Fallout 4 VR files: {str(e)}", [])
            
            # Check for missing files
            missing_files = []
            for req_file in required_files:
                if req_file.lower() not in existing_files_lower:
                    missing_files.append(req_file)
            
            if missing_files:
                logging.warning(f"Fallout 4 VR missing {len(missing_files)} files: {missing_files}")
                return (False, "Fallout 4 VR files incomplete. Please redownload.", missing_files)
            
            # All required files present
            logging.info(f"Fallout 4 VR installation validated at {data_dir}")
            return (True, "Fallout 4 VR: Ready for installation", [])
            
        except Exception as e:
            logging.error(f"Error validating Fallout 4 VR files: {e}")
            return (False, f"Error validating Fallout 4 VR files: {str(e)}", [])

    def reorder_status_labels(self):
        """Reorder status labels: green on top, red/warning below, disk space at bottom.
        
        This function unpacks and repacks status labels in the correct order.
        """
        try:
            # Skip if in update mode - no need to reorder
            if self.is_update_detected and self.update_mode:
                return
            
            # Check if content_frame exists
            if not hasattr(self, 'content_frame') or not self.content_frame or not self.content_frame.winfo_exists():
                logging.warning("content_frame doesn't exist, skipping reorder_status_labels")
                return
            
            # Temporarily hide the content frame to prevent flickering during reorder
            self.content_frame.pack_propagate(False)
            
            # Collect all status labels with their current state
            labels_info = []
            
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                text = self.dlc_status_label.cget('text')
                fg = self.dlc_status_label.cget('fg').lower()
                if text and text.strip():
                    labels_info.append(('dlc', self.dlc_status_label, text, fg))
                    
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label and self.f4vr_status_label.winfo_exists():
                text = self.f4vr_status_label.cget('text')
                fg = self.f4vr_status_label.cget('fg').lower()
                if text and text.strip():
                    labels_info.append(('f4vr', self.f4vr_status_label, text, fg))
                    
            if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
                text = self.london_status_label.cget('text')
                fg = self.london_status_label.cget('fg').lower()
                if text and text.strip():
                    labels_info.append(('london', self.london_status_label, text, fg))
                    
            if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                text = self.message_label.cget('text')
                fg = self.message_label.cget('fg').lower()
                if text and text.strip():
                    labels_info.append(('message', self.message_label, text, fg))
            
            # Categorize labels by name for proper ordering
            # Green/Orange = ready (treat orange same as green so it stays in place), Red = not ready/error
            green_labels_by_name = {}
            red_labels_by_name = {}
            
            for name, label, text, fg in labels_info:
                if fg in ['#00ff00', 'green', '#ffa500', 'orange']:
                    # Treat orange same as green - both are "ready" states
                    green_labels_by_name[name] = label
                elif fg in ['#ff6666', 'red']:
                    red_labels_by_name[name] = label
            
            # Define the display order: London first, then F4VR, then DLC
            label_order = ['london', 'f4vr', 'dlc', 'message']
            green_labels = [green_labels_by_name[name] for name in label_order if name in green_labels_by_name]
            orange_labels = []  # No longer used - orange treated as green
            red_labels = [red_labels_by_name[name] for name in label_order if name in red_labels_by_name]
            
            # Unpack ALL children of content_frame temporarily
            all_children = list(self.content_frame.winfo_children())
            pack_info_map = {}
            
            for child in all_children:
                try:
                    if child.winfo_ismapped():
                        pack_info_map[child] = child.pack_info()
                except:
                    pass
                child.pack_forget()
            
            # Now repack in the correct order:
            # 1. Green status labels first (includes orange - both are "ready" states)
            for label in green_labels:
                label.pack(pady=self.get_scaled_value(2), fill="x")
            
            # 2. Red status labels (not ready/errors)
            for label in red_labels:
                label.pack(pady=self.get_scaled_value(2), fill="x")
            
            # 4. Disk space label
            if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                disk_text = self.disk_space_label.cget('text')
                if disk_text and disk_text.strip():
                    self.disk_space_label.pack(pady=self.get_scaled_value(2), fill="x")
            
            # 5. Message label (if it has content and wasn't already packed)
            if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                if self.message_label not in green_labels and self.message_label not in orange_labels and self.message_label not in red_labels:
                    msg_text = self.message_label.cget('text')
                    if msg_text and msg_text.strip():
                        self.message_label.pack(pady=self.get_scaled_value(2), fill="x")
            
            # Build list of status/disk labels to exclude from general repacking
            status_and_disk = [self.dlc_status_label if hasattr(self, 'dlc_status_label') else None,
                              self.f4vr_status_label if hasattr(self, 'f4vr_status_label') else None,
                              self.london_status_label if hasattr(self, 'london_status_label') else None,
                              self.message_label if hasattr(self, 'message_label') else None,
                              self.disk_space_label if hasattr(self, 'disk_space_label') else None]
            
            # Build list of widgets that should NOT be auto-packed (handled separately)
            special_widgets = [self.atkins_label if hasattr(self, 'atkins_label') else None]
            
            # 6. Repack all other widgets - use saved info if available, otherwise use defaults
            for child in all_children:
                if child not in status_and_disk and child not in special_widgets:
                    if child in pack_info_map:
                        info = pack_info_map[child]
                        child.pack(
                            pady=info.get('pady', 0),
                            padx=info.get('padx', 0),
                            fill=info.get('fill', 'none'),
                            expand=info.get('expand', False),
                            side=info.get('side', 'top'),
                            anchor=info.get('anchor', 'center')
                        )
                    elif child.winfo_exists():
                        # Widget wasn't mapped before but exists - pack with defaults
                        # This ensures widgets that were created but not yet packed don't get lost
                        child.pack(pady=self.get_scaled_value(5))
            
            # 7. FAILSAFE: Ensure critical widgets are always visible in fresh install mode
            self._ensure_critical_widgets_packed()
            
            # Re-enable pack propagation and force update
            self.content_frame.pack_propagate(True)
            self.content_frame.update_idletasks()
                
        except Exception as e:
            logging.warning(f"Error reordering status labels: {e}")
            # Ensure pack_propagate is restored even on error
            try:
                if hasattr(self, 'content_frame') and self.content_frame and self.content_frame.winfo_exists():
                    self.content_frame.pack_propagate(True)
            except:
                pass

    def _ensure_critical_widgets_packed(self):
        """Failsafe to ensure critical widgets are always packed in fresh install mode"""
        try:
            if self.is_update_detected and self.update_mode:
                return  # Don't apply in update mode
            
            # Check if London widgets should be visible
            if hasattr(self, 'london_label') and self.london_label and self.london_label.winfo_exists():
                if not self.london_label.winfo_ismapped():
                    self.london_label.pack(pady=(10, 5))
            if hasattr(self, 'london_data_entry') and self.london_data_entry and self.london_data_entry.winfo_exists():
                if not self.london_data_entry.winfo_ismapped():
                    self.london_data_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'london_browse_button') and self.london_browse_button and self.london_browse_button.winfo_exists():
                if not self.london_browse_button.winfo_ismapped():
                    self.london_browse_button.pack()
            
            # Check if F4VR widgets should be visible
            if hasattr(self, 'f4vr_label') and self.f4vr_label and self.f4vr_label.winfo_exists():
                if not self.f4vr_label.winfo_ismapped():
                    self.f4vr_label.pack(pady=(10, 5))
            if hasattr(self, 'f4vr_entry') and self.f4vr_entry and self.f4vr_entry.winfo_exists():
                if not self.f4vr_entry.winfo_ismapped():
                    self.f4vr_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'f4vr_browse_button') and self.f4vr_browse_button and self.f4vr_browse_button.winfo_exists():
                if not self.f4vr_browse_button.winfo_ismapped():
                    self.f4vr_browse_button.pack()
            
            # Check if installation directory container should be visible
            if hasattr(self, 'installation_dir_container') and self.installation_dir_container and self.installation_dir_container.winfo_exists():
                if not self.installation_dir_container.winfo_ismapped():
                    self.installation_dir_container.pack(pady=(0, 2))
            
            # Check if instruction label should be visible
            if hasattr(self, 'instruction_label') and self.instruction_label and self.instruction_label.winfo_exists():
                if not self.instruction_label.winfo_ismapped():
                    self.instruction_label.pack(pady=(40, 5))
            
            # Check if install button should be visible
            if hasattr(self, 'install_button') and self.install_button and self.install_button.winfo_exists():
                if not self.install_button.winfo_ismapped():
                    self.install_button.pack(pady=(2, 20))
            
            # Don't show Atkins on fresh install screen - only show in update mode
                    
        except Exception as e:
            logging.warning(f"Error in _ensure_critical_widgets_packed: {e}")

    def show_f4_dlc_widgets(self):
        """Show Fallout 4 DLC path widgets (when DLC not found in London or F4VR paths)"""
        try:
            if hasattr(self, 'f4_label') and hasattr(self, 'content_frame'):
                # First, temporarily hide the installation directory and instruction widgets
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack_forget()
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack_forget()
                if hasattr(self, 'install_button'):
                    self.install_button.pack_forget()
                
                # Pack F4 DLC widgets with minimal padding
                self.f4_label.pack(pady=(10, 5))
                self.f4_entry.pack(pady=self.get_scaled_value(5))
                self.f4_browse_button.pack()
                
                # Re-pack installation directory widgets after F4 DLC widgets
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack(pady=(0, 2))
                
                # Re-pack instruction label and install button at the end
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack(pady=(40, 5))
                if hasattr(self, 'install_button'):
                    self.install_button.pack(pady=(2, 20))
                    
                logging.info("Showed F4 DLC widgets - DLC not found in London or F4VR paths")
        except Exception as e:
            logging.warning(f"Error showing F4 DLC widgets: {e}")

    def show_london_widgets(self):
        """Show London data widgets"""
        try:
            if hasattr(self, 'london_label') and hasattr(self, 'content_frame'):
                # First, temporarily hide the installation directory and instruction widgets
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack_forget()
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack_forget()
                if hasattr(self, 'install_button'):
                    self.install_button.pack_forget()
                
                # Pack London widgets with minimal padding
                self.london_label.pack(pady=(5, 2))
                self.london_data_entry.pack(pady=2)
                self.london_browse_button.pack(pady=2)
                
                # Re-pack installation directory widgets after London widgets with reduced padding
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack(pady=(0, 2))
                
                # Re-pack instruction label and install button at the end with reduced padding
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack(pady=(40, 5))
                if hasattr(self, 'install_button'):
                    self.install_button.pack(pady=(2, 20))
                    
        except Exception as e:
            logging.warning(f"Error showing London widgets: {e}")
            # Fallback: recreate widgets if there's an issue
            self.recreate_london_widgets()

    def recreate_london_widgets(self):
        """Recreate London widgets if there's an issue with packing"""
        try:
            if hasattr(self, 'content_frame'):
                bg_color = "#1e1e1e"
                fg_color = "#ffffff"
                entry_fg_color = "#000000"
                entry_bg_color = "#d3d3d3"
                accent_color = "#0078d7"
                
                # Temporarily hide instruction label and install button
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack_forget()
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack_forget()
                if hasattr(self, 'install_button'):
                    self.install_button.pack_forget()
                
                # Recreate widgets
                self.london_label = tk.Label(self.content_frame, text="Fallout London Location", font=self.bold_font, bg=bg_color, fg=fg_color)
                self.london_data_entry = tk.Entry(self.content_frame, textvariable=self.london_data_path, width=40, font=self.regular_font, bg=entry_bg_color, fg=entry_fg_color, insertbackground=fg_color, bd=0)
                self.london_browse_button = tk.Button(self.content_frame, text="Browse", command=lambda: self.browse_path(self.london_data_path), font=self.regular_font, bg=accent_color, fg=fg_color, bd=0, relief="flat", activebackground="#005ba1", padx=10, pady=5)
                
                # Pack them with minimal padding
                self.london_label.pack(pady=(5, 2))
                self.london_data_entry.pack(pady=2)
                self.london_browse_button.pack(pady=2)
                
                # Re-pack installation directory with reduced padding
                if hasattr(self, 'installation_dir_container'):
                    self.installation_dir_container.pack(pady=(0, 2))
                
                # Re-pack instruction label and install button at the end with reduced padding
                if hasattr(self, 'instruction_label'):
                    self.instruction_label.pack(pady=(40, 5))
                if hasattr(self, 'install_button'):
                    self.install_button.pack(pady=(2, 20))
                
                logging.info("Recreated London widgets successfully")
        except Exception as e:
            logging.error(f"Failed to recreate London widgets: {e}")

    def hide_input_widgets_only(self):
        """Hide only the input widgets while keeping the window structure"""
        try:
            # Hide status labels
            if hasattr(self, 'f4_status_label') and self.f4_status_label:
                self.f4_status_label.pack_forget()
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label:
                self.f4vr_status_label.pack_forget()
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label:
                self.dlc_status_label.pack_forget()
            if hasattr(self, 'london_status_label') and self.london_status_label:
                self.london_status_label.pack_forget()
            if hasattr(self, 'disk_space_label') and self.disk_space_label:
                self.disk_space_label.pack_forget()
            
            # Hide all Entry widgets and buttons, but preserve logo
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for subchild in child.winfo_children():
                                # Check if this is the logo label
                                if isinstance(subchild, tk.Label) and hasattr(subchild, 'cget'):
                                    try:
                                        # Skip if this label contains the logo image
                                        if subchild.cget('image') and self.logo and str(subchild.cget('image')) == str(self.logo):
                                            continue
                                    except:
                                        pass
                                
                                if isinstance(subchild, (tk.Entry, tk.Button)):
                                    subchild.pack_forget()
                                elif isinstance(subchild, tk.Label):
                                    # Hide labels except logo, title, and message label
                                    try:
                                        label_text = str(subchild.cget("text"))
                                        if (subchild != self.message_label and 
                                            "Fallout: London VR Installation" not in label_text and
                                            subchild.cget('image') != str(self.logo)):
                                            subchild.pack_forget()
                                    except:
                                        pass
            
            # Hide install button
            if hasattr(self, 'install_button') and self.install_button:
                self.install_button.pack_forget()
            
            # Update message
            if self.message_label:
                self.message_label.config(text="Starting installation", justify="center", anchor="center")
            
            logging.info("Input widgets hidden successfully, logo preserved")
        except Exception as e:
            logging.warning(f"Error hiding input widgets: {e}")

    def initialize_and_start_slideshow(self):
        """Initialize and start the slideshow after install button is clicked"""
        try:
            logging.info("=== INITIALIZING SLIDESHOW (FRESH INSTALL) ===")
            
            # Create slideshow if not already created
            if not hasattr(self, 'slideshow') or self.slideshow is None:
                # Find the content frame or main container
                content_parent = None
                
                # First try to find the main container (should exist)
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Frame):
                        # Check if this is the main container frame
                        for child in widget.winfo_children():
                            if isinstance(child, tk.Frame):
                                # This is likely our content frame
                                content_parent = child
                                logging.info(f"Found content frame: {content_parent}")
                                break
                        if content_parent:
                            break
                
                # If we didn't find a nested frame, use the first frame we find
                if not content_parent:
                    for widget in self.root.winfo_children():
                        if isinstance(widget, tk.Frame):
                            content_parent = widget
                            logging.info(f"Using first frame found: {content_parent}")
                            break
                
                # Final fallback to root
                if not content_parent:
                    content_parent = self.root
                    logging.warning("No frame found, using root window for slideshow")
                
                logging.info(f"Parent selected for slideshow: {content_parent}")
                
                # Create the slideshow instance
                self.slideshow = SlideshowFrame(content_parent, self)
                logging.info("Slideshow instance created")
            else:
                logging.info("Slideshow already exists")
            
            # Start the slideshow
            if self.slideshow:
                self.slideshow.start_slideshow()
                logging.info("Slideshow start method called")
            else:
                logging.error("Slideshow instance is None!")
            
        except Exception as e:
            logging.error(f"Failed to start slideshow: {e}")
            import traceback
            logging.error(traceback.format_exc())
            # Non-critical feature, continue without slideshow

    def check_london_installation(self, dlc_path):
        """Check if London files are present in the DLC path root or Data folder only"""
        self.london_installed = False
        
        dlc_dir = Path(dlc_path)
        
        if not dlc_dir.exists():
            return
        
        london_esm_found = False
        
        # Check root directory
        london_esm_found = self._check_london_in_directory(dlc_dir)
        
        # If not found in root, check Data subdirectory
        if not london_esm_found:
            data_dir = dlc_dir / "Data"
            if data_dir.exists():
                london_esm_found = self._check_london_in_directory(data_dir)
        
        logging.info(f"Checking for London in {dlc_path}")
        logging.info(f"LondonWorldSpace ESM found: {london_esm_found}")
        
        if london_esm_found:
            self.london_installed = True
            # Only hide London widgets in update mode, not fresh install
            if self.is_update_detected and self.update_mode:
                self.hide_london_widgets()
            else:
                # In fresh install mode, set the path but keep widgets visible
                self.london_data_path.set(dlc_path)
                if hasattr(self, 'london_status_label') and self.london_status_label.winfo_exists():
                    # Detect version by checking for Textures14.ba2
                    is_103 = self._check_london_version_103(dlc_path)
                    if is_103:
                        self.root.after(0, lambda: self.london_status_label.config(text="Fallout: London 1.03 Rabbit & Pork: Ready for installation", fg="#00ff00"))
                    else:
                        self.root.after(0, lambda: self.london_status_label.config(text="Fallout: London 1.02: Ready for installation", fg="#ffa500"))
            logging.info("Fallout: London detected - LondonWorldSpace file found.")
        else:
            self.london_installed = False
            self.london_data_path.set("")
            self.show_london_widgets()
            logging.info("Fallout: London NOT found - no LondonWorldSpace files detected")
            self.search_for_london_installation()
        
        self.check_version(dlc_path)
        # Reorder status labels after London check
        self.root.after(100, self.reorder_status_labels)

    def _check_london_in_directory(self, directory):
        """Check a single directory for London files (non-recursive)"""
        try:
            # Only check files in this specific directory
            for file in directory.iterdir():
                if file.is_file():
                    file_lower = file.name.lower()
                    if file_lower.startswith("londonworldspace") and file_lower.endswith(".esm"):
                        return True
            return False
        except Exception as e:
            logging.error(f"Error checking directory {directory}: {e}")
            return False

    def search_for_london_installation(self):
        """Search for standalone London installation - only check root and Data folders"""
        # Get all active drives
        active_drives = self.get_all_active_drives()
        
        # Build search paths for London installations
        search_paths = []
        
        for drive in active_drives:
            search_paths.extend([
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common",
                f"{drive}\\Steam\\steamapps\\common",
                f"{drive}\\SteamLibrary\\steamapps\\common",
                f"{drive}\\GOG Games",
                f"{drive}\\Program Files (x86)\\GOG Galaxy\\Games",
                f"{drive}\\Games",
                f"{drive}\\Downloads",
                f"{drive}\\Program Files\\Epic Games",
                f"{drive}\\Epic Games"
            ])
        
        # Remove duplicates
        unique_paths = []
        for path in search_paths:
            if path not in unique_paths:
                unique_paths.append(path)
        
        for base_path in unique_paths:
            if os.path.exists(base_path):
                try:
                    # Only check immediate subdirectories, not recursive
                    subdirs = [d for d in os.listdir(base_path) 
                              if os.path.isdir(os.path.join(base_path, d))]
                    
                    for folder in subdirs:
                        # Check folders starting with or containing 'fallout'
                        if folder.lower().startswith("fallout") or "fallout" in folder.lower():
                            possible_path = os.path.join(base_path, folder)
                            
                            # Check root directory for London files
                            if self._check_london_files_in_path(possible_path):
                                self.london_data_path.set(possible_path)
                                if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
                                    is_103 = self._check_london_version_103(possible_path)
                                    if is_103:
                                        self.root.after(0, lambda: self.london_status_label.config(
                                            text="Fallout: London 1.03 Rabbit & Pork: Ready for installation", fg="#00ff00"))
                                    else:
                                        self.root.after(0, lambda: self.london_status_label.config(
                                            text="Fallout: London 1.02: Ready for installation", fg="#ffa500"))
                                logging.info(f"Fallout: London files found at {possible_path}")
                                return
                            
                            # Check Data subdirectory
                            data_path = os.path.join(possible_path, "Data")
                            if os.path.exists(data_path) and self._check_london_files_in_path(data_path):
                                self.london_data_path.set(possible_path)  # Set to root, not Data
                                if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
                                    is_103 = self._check_london_version_103(possible_path)
                                    if is_103:
                                        self.root.after(0, lambda: self.london_status_label.config(
                                            text="Fallout: London 1.03 Rabbit & Pork: Ready for installation", fg="#00ff00"))
                                    else:
                                        self.root.after(0, lambda: self.london_status_label.config(
                                            text="Fallout: London 1.02: Ready for installation", fg="#ffa500"))
                                logging.info(f"Fallout: London files found in Data folder at {possible_path}")
                                return
                                
                except Exception as e:
                    logging.warning(f"Error scanning directory {base_path}: {e}")
                    continue
        
        if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
            self.root.after(0, lambda: self.london_status_label.config(
                text="Fallout: London: Invalid path", fg="#ff6666"))
        logging.info("No Fallout: London files found in common locations")

    def _check_london_files_in_path(self, path):
        """Check if London files exist in a specific path (non-recursive)"""
        try:
            # Only check files in this specific directory
            files = os.listdir(path)
            for file in files:
                if os.path.isfile(os.path.join(path, file)):
                    if file.lower().startswith("londonworldspace"):
                        return True
            return False
        except Exception as e:
            logging.warning(f"Error checking path {path}: {e}")
            return False

    def _check_london_version_103(self, path):
        """Check if London installation is version 1.03 by looking for Textures14.ba2"""
        try:
            # Check root directory
            textures14_root = os.path.join(path, "LondonWorldSpace - Textures14.ba2")
            if os.path.exists(textures14_root):
                return True
            
            # Check Data subdirectory
            data_path = os.path.join(path, "Data")
            if os.path.exists(data_path):
                textures14_data = os.path.join(data_path, "LondonWorldSpace - Textures14.ba2")
                if os.path.exists(textures14_data):
                    return True
            
            return False
        except Exception as e:
            logging.warning(f"Error checking London version at {path}: {e}")
            return False

    def get_all_active_drives(self) -> list[str]:
        """Get all active drive letters on the system"""
        active_drives = []
        
        # Method 1: Using psutil (more reliable)
        try:
            partitions = psutil.disk_partitions()
            for partition in partitions:
                # Get drive letter (e.g., "C:\" -> "C:")
                drive_letter = partition.device.rstrip('\\')
                if drive_letter not in active_drives:
                    active_drives.append(drive_letter)
            logging.info(f"Detected drives via psutil: {active_drives}")
        except Exception as e:
            logging.warning(f"psutil drive detection failed: {e}")
        
        # Method 2: Fallback using Windows API
        if not active_drives:
            try:
                import ctypes
                drives = []
                bitmask = ctypes.windll.kernel32.GetLogicalDrives()
                for letter in string.ascii_uppercase:
                    if bitmask & 1:
                        drive = f"{letter}:"
                        # Verify drive is accessible
                        try:
                            if os.path.exists(f"{drive}\\"):
                                drives.append(drive)
                        except:
                            pass
                    bitmask >>= 1
                active_drives = drives
                logging.info(f"Detected drives via Windows API: {active_drives}")
            except Exception as e:
                logging.warning(f"Windows API drive detection failed: {e}")
        
        # Method 3: Final fallback
        if not active_drives:
            for letter in string.ascii_uppercase:
                drive = f"{letter}:"
                if os.path.exists(f"{drive}\\"):
                    active_drives.append(drive)
            logging.info(f"Detected drives via fallback method: {active_drives}")
        
        return active_drives

    def scan_for_fallout_installations(self, search_paths):
        """Scan paths for DLC files and F4VR - only check root and Data folders"""
        valid_dlc_paths = []
        valid_f4vr_paths = []
        
        for base_path in search_paths:
            if not os.path.exists(base_path):
                continue
            
            try:
                # Get immediate subdirectories only
                directories = [d for d in os.listdir(base_path) 
                              if os.path.isdir(os.path.join(base_path, d))]
                
                for folder in directories:
                    # Check folders starting with or containing 'fallout'
                    if folder.lower().startswith("fallout") or "fallout" in folder.lower():
                        game_path = os.path.join(base_path, folder)
                        
                        # Check for Fallout4.exe in root
                        f4_exe = os.path.join(game_path, "Fallout4.exe")
                        if os.path.exists(f4_exe):
                            # Check for DLC in root or Data folder
                            has_dlc = self._check_dlc_presence(game_path)
                            if has_dlc:
                                has_london = self._check_london_presence(game_path)
                                is_next_gen = self._check_if_next_gen(game_path)
                                is_gog = "gog" in base_path.lower()
                                valid_dlc_paths.append((game_path, has_london, is_next_gen, is_gog))
                                logging.info(f"Found DLC path: {game_path} (London: {has_london}, NG: {is_next_gen}, GOG: {is_gog})")
                        
                        # F4VR check
                        f4vr_exe = os.path.join(game_path, "Fallout4VR.exe")
                        if os.path.exists(f4vr_exe):
                            version, is_next_gen = self.get_fallout4_version(f4vr_exe)
                            is_gog = "gog" in base_path.lower()
                            valid_f4vr_paths.append((game_path, is_next_gen, is_gog))
                            logging.info(f"Found Fallout 4 VR: {game_path} (NG: {is_next_gen}, GOG: {is_gog})")
                            
            except Exception as e:
                logging.warning(f"Error scanning directory {base_path}: {e}")
                continue
        
        self.select_best_installations(valid_dlc_paths, valid_f4vr_paths)

    def _check_dlc_presence(self, game_path):
        """Check if DLC files are present in root or Data folder"""
        required_esm = ["DLCRobot.esm", "DLCworkshop01.esm", "DLCCoast.esm", 
                        "DLCworkshop02.esm", "DLCworkshop03.esm", "DLCNukaWorld.esm"]
        
        # Check root directory
        root_files = set()
        try:
            root_files = {f for f in os.listdir(game_path) 
                         if os.path.isfile(os.path.join(game_path, f))}
        except:
            pass
        
        # Check Data directory
        data_files = set()
        data_path = os.path.join(game_path, "Data")
        if os.path.exists(data_path):
            try:
                data_files = {f for f in os.listdir(data_path) 
                             if os.path.isfile(os.path.join(data_path, f))}
            except:
                pass
        
        all_files = root_files | data_files
        
        # Check if all required DLC files are present
        return all(esm in all_files for esm in required_esm)

    def _check_london_presence(self, game_path):
        """Check if London files are present in root or Data folder"""
        # Check root
        if self._check_london_files_in_path(game_path):
            return True
        
        # Check Data folder
        data_path = os.path.join(game_path, "Data")
        if os.path.exists(data_path):
            return self._check_london_files_in_path(data_path)
        
        return False

    def _check_if_next_gen(self, game_path):
        """Check if installation has Next-Gen DLC files"""
        # Check root directory for BA2 files
        try:
            for file in os.listdir(game_path):
                if file.endswith('.ba2'):
                    file_path = os.path.join(game_path, file)
                    header = BA2Header.from_file(file_path)
                    if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                        return True
        except:
            pass
        
        # Check Data directory
        data_path = os.path.join(game_path, "Data")
        if os.path.exists(data_path):
            try:
                for file in os.listdir(data_path):
                    if file.endswith('.ba2'):
                        file_path = os.path.join(data_path, file)
                        header = BA2Header.from_file(file_path)
                        if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                            return True
            except:
                pass
        
        return False

    def select_best_installations(self, valid_f4_paths, valid_f4vr_paths):
        """Select the best Fallout 4 and Fallout 4 VR installations"""
        # Fallout 4 selection logic
        selected_f4_path = None
        if valid_f4_paths:
            # Priority 1: Pre-Next-Gen with LondonWorldSpace.esm
            pre_ng_london_paths = [p for p, has_london, is_next_gen, _ in valid_f4_paths if has_london and not is_next_gen]
            if pre_ng_london_paths:
                selected_f4_path = pre_ng_london_paths[0]
                logging.info(f"Selected pre-Next-Gen Fallout 4 path with LondonWorldSpace.esm: {selected_f4_path}")
            else:
                # Priority 2: Any path with LondonWorldSpace.esm
                london_paths = [p for p, has_london, _, _ in valid_f4_paths if has_london]
                if london_paths:
                    selected_f4_path = london_paths[0]
                    logging.info(f"Selected Fallout 4 path with LondonWorldSpace.esm: {selected_f4_path}")
                else:
                    # Priority 3: Pre-Next-Gen or GOG
                    pre_ng_or_gog_paths = [p for p, _, is_next_gen, is_gog in valid_f4_paths if not is_next_gen or is_gog]
                    if pre_ng_or_gog_paths:
                        selected_f4_path = pre_ng_or_gog_paths[0]
                        logging.info(f"Selected Fallout 4 path (pre-NG or GOG): {selected_f4_path}")
                    else:
                        # Fallback: First valid path
                        selected_f4_path = valid_f4_paths[0][0]
                        logging.info(f"Selected first valid Fallout 4 path: {selected_f4_path}")
            
            self.f4_path.set(selected_f4_path)
            self.check_london_installation(selected_f4_path)
        else:
            # Show F4 not found status - use dlc_status_label since f4_status_label doesn't exist
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                self.root.after(0, lambda: self.dlc_status_label.config(text="Fallout 4: Not detected", fg="#ff6666"))
            if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                self.root.after(0, lambda: self.message_label.config(text="Please select the Fallout 4 installation folder.", fg="#ffffff"))
            logging.info("Fallout 4 not detected in common locations or missing Fallout4.exe")
            
            # Still check for London installation independently
            self.london_installed = False
            self.london_data_path.set("")
            self.show_london_widgets()
            self.search_for_london_installation()
            
            # Update London status based on search results
            if self.london_data_path.get():
                self.root.after(0, lambda: self.london_status_label.config(text="Fallout: London: Download found", fg="#00ff00") if self.london_status_label.winfo_exists() else None)
            else:
                self.root.after(0, lambda: self.london_status_label.config(text="Fallout: London: Not detected. Select the path to the files you downloaded from GOG.", fg="#ff6666") if self.london_status_label.winfo_exists() else None)

        # Fallout 4 VR selection logic
        selected_f4vr_path = None
        if valid_f4vr_paths:
            # Priority: Pre-Next-Gen or GOG
            pre_ng_or_gog_paths = [p for p, is_next_gen, is_gog in valid_f4vr_paths if not is_next_gen or is_gog]
            if pre_ng_or_gog_paths:
                selected_f4vr_path = pre_ng_or_gog_paths[0]
                logging.info(f"Selected Fallout 4 VR path (pre-NG or GOG): {selected_f4vr_path}")
            else:
                # Fallback: First valid path
                selected_f4vr_path = valid_f4vr_paths[0][0]
                logging.info(f"Selected first valid Fallout 4 VR path: {selected_f4vr_path}")
            
            self.f4vr_path.set(selected_f4vr_path)
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label and self.f4vr_status_label.winfo_exists():
                self.root.after(0, lambda: self.f4vr_status_label.config(text="Fallout 4 VR: Ready for installation", fg="#00ff00"))
                logging.info("Updated f4vr_status_label to 'Detected' in detect_paths")
        else:
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label and self.f4vr_status_label.winfo_exists():
                self.root.after(0, lambda: self.f4vr_status_label.config(text="Fallout 4 VR: Not detected", fg="#ff6666"))
                logging.info("Updated f4vr_status_label to 'Not found' in detect_paths")
            logging.info("Fallout 4 VR not detected in common locations or missing Fallout4VR.exe")
        
        # Check DLC status independently (can be in either F4 or F4VR)
        self.check_dlc_status_independent()

    def on_installation_path_changed(self):
        """Handle user changes to installation directory path"""
        # If user manually changes the installation path, restart in fresh install mode
        if self.is_update_detected and self.mo2_path.get() != self.detected_install_path:
            logging.info("User changed installation path - restarting in Fresh Install mode")
            self.restart_in_fresh_install_mode()
            return
        
        # Update disk space regardless
        self.update_disk_space()
    
    def restart_in_fresh_install_mode(self):
        """Switch to fresh install mode WITHOUT restarting the process.
        
        We cannot restart the process because PyInstaller creates a new _MEI
        temp folder each time, which loses access to bundled assets and causes
        'failed to remove temp directory' errors.
        
        Instead, we reset the UI state and re-run detection in the same process.
        """
        try:
            logging.info("Switching to fresh install mode (in-process, no restart)")
            
            # Get current install path to preserve it
            install_path = self.mo2_path.get()
            
            # Reset all state flags
            self.is_update_detected = False
            self.update_mode = False
            self.skip_update_detection = True  # Prevent re-detecting as update
            self.london_installed = False
            self.missing_dlc = []
            
            # Destroy update-specific widgets completely (not just hide)
            if hasattr(self, 'update_button') and self.update_button:
                try:
                    self.update_button.destroy()
                    self.update_button = None
                except:
                    pass
            if hasattr(self, 'update_instruction_label') and self.update_instruction_label:
                try:
                    self.update_instruction_label.destroy()
                    self.update_instruction_label = None
                except:
                    pass
            
            # Update install button text
            if hasattr(self, 'install_button') and self.install_button and self.install_button.winfo_exists():
                self.install_button.config(text="Install")
            
            # Clear installation status
            if hasattr(self, 'installation_status_label') and self.installation_status_label and self.installation_status_label.winfo_exists():
                self.installation_status_label.config(text="")
            
            # Restore UI for fresh install
            self._restore_ui_for_fresh_install()
            
            # Run fresh install detection after a short delay
            self.root.after(200, self._run_fresh_install_detection)
            
            logging.info(f"Switched to fresh install mode with path: {install_path}")
            
        except Exception as e:
            logging.error(f"Failed to switch to fresh install mode: {e}")
    
    def _run_fresh_install_detection(self):
        """Run detection for fresh install mode (London, F4VR, DLC)"""
        try:
            # Reset detection state
            self.london_installed = False
            self.missing_dlc = []
            # Clear paths to allow fresh detection
            self.f4_path.set("")
            self.f4vr_path.set("")
            self.london_data_path.set("")
            
            # Clear status labels
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label.winfo_exists():
                self.f4vr_status_label.config(text="Fallout 4 VR: Detecting...", fg="#ffffff")
            if hasattr(self, 'london_status_label') and self.london_status_label.winfo_exists():
                self.london_status_label.config(text="Fallout: London: Detecting...", fg="#ffffff")
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label.winfo_exists():
                self.dlc_status_label.config(text="DLC: Detecting...", fg="#ffffff")
            
            # Run the original robust detection in background thread
            # skip_update_detection is already set to True, so detect_paths will skip update check
            def detect_thread():
                self.detect_paths()
            
            threading.Thread(target=detect_thread, daemon=True).start()
            
            logging.info("Started fresh install detection using original detect_paths()")
            
        except Exception as e:
            logging.error(f"Error running fresh install detection: {e}")

    def _restore_ui_for_fresh_install(self):
        """Restore UI for fresh install: hide atkins image and restore original order"""
        try:
            # Check if content_frame exists
            if not hasattr(self, 'content_frame') or not self.content_frame or not self.content_frame.winfo_exists():
                logging.warning("content_frame doesn't exist, cannot restore UI for fresh install")
                return
            
            # Unpack all widgets from content_frame to reorder them back to original
            for widget in self.content_frame.winfo_children():
                widget.pack_forget()
            
            # Hide update-specific widgets if they exist
            if hasattr(self, 'update_instruction_label') and self.update_instruction_label and self.update_instruction_label.winfo_exists():
                self.update_instruction_label.pack_forget()
            if hasattr(self, 'update_button') and self.update_button and self.update_button.winfo_exists():
                self.update_button.pack_forget()
            
            bg_color = "#1e1e1e"
            fg_color = "#ffffff"
            entry_fg_color = "#000000"
            entry_bg_color = "#d3d3d3"
            accent_color = "#0078d7"
            
            # Re-pack in new order for fresh install: London path first, then F4VR, then F4 DLC (conditional)
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                self.dlc_status_label.pack(pady=self.get_scaled_value(2), fill="x")
            if hasattr(self, 'f4vr_status_label') and self.f4vr_status_label and self.f4vr_status_label.winfo_exists():
                self.f4vr_status_label.pack(pady=self.get_scaled_value(2), fill="x")
            if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
                self.london_status_label.pack(pady=self.get_scaled_value(2), fill="x")
            if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                self.disk_space_label.pack(pady=self.get_scaled_value(2), fill="x")
            if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                self.message_label.pack(pady=self.get_scaled_value(2), fill="x")
            
            # Hide installation status label (welcome message) for fresh install
            if hasattr(self, 'installation_status_label') and self.installation_status_label and self.installation_status_label.winfo_exists():
                self.installation_status_label.pack_forget()
            
            # Hide atkins image
            if hasattr(self, 'atkins_label') and self.atkins_label and self.atkins_label.winfo_exists():
                self.atkins_label.pack_forget()
            
            # Show London path section FIRST
            if hasattr(self, 'london_label') and self.london_label and self.london_label.winfo_exists():
                self.london_label.pack(pady=(10, 5))
            if hasattr(self, 'london_data_entry') and self.london_data_entry and self.london_data_entry.winfo_exists():
                self.london_data_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'london_browse_button') and self.london_browse_button and self.london_browse_button.winfo_exists():
                self.london_browse_button.pack()
            
            # Show F4VR path section SECOND
            if hasattr(self, 'f4vr_label') and self.f4vr_label and self.f4vr_label.winfo_exists():
                self.f4vr_label.pack(pady=(10, 5))
            if hasattr(self, 'f4vr_entry') and self.f4vr_entry and self.f4vr_entry.winfo_exists():
                self.f4vr_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'f4vr_browse_button') and self.f4vr_browse_button and self.f4vr_browse_button.winfo_exists():
                self.f4vr_browse_button.pack()
            
            # F4 DLC path is NOT shown here - it will be shown conditionally by show_f4_dlc_widgets() if needed
            
            # Show installation directory at the end
            if hasattr(self, 'installation_dir_container') and self.installation_dir_container and self.installation_dir_container.winfo_exists():
                # Reset the label text back to fresh install mode
                if hasattr(self, 'installation_dir_label') and self.installation_dir_label and self.installation_dir_label.winfo_exists():
                    self.installation_dir_label.config(text="Installation Location")
                self.installation_dir_container.pack(pady=(0, 2))
            
            # Reset header label back to "Installation"
            if hasattr(self, 'header_label') and self.header_label and self.header_label.winfo_exists():
                self.header_label.config(text="Fallout: London VR Installation")
            
            # Show instruction label and install button in content_frame (same position as update mode)
            if hasattr(self, 'instruction_label') and self.instruction_label and self.instruction_label.winfo_exists():
                self.instruction_label.config(text="Press Install or Enter to continue")
                self.instruction_label.pack(pady=(40, 5))
            if hasattr(self, 'install_button') and self.install_button and self.install_button.winfo_exists():
                self.install_button.config(text="Install")
                self.install_button.pack(pady=(2, 20))
            
            # Rebind Enter key for fresh install mode
            self.root.bind('<Return>', lambda event: self.validate_and_install())
            
            # Clean up any bottom frame installation widgets
            if hasattr(self, 'installation_bottom_container') and self.installation_bottom_container and self.installation_bottom_container.winfo_exists():
                self.installation_bottom_container.pack_forget()
            
            logging.info("Restored UI for fresh install mode")
        except Exception as e:
            logging.warning(f"Error restoring UI for fresh install: {e}")
    
    def _hide_dlc_and_f4vr_fields(self):
        """Hide London, F4VR and F4 DLC path fields for update mode"""
        try:
            # Hide London path
            if hasattr(self, 'london_label'):
                self.london_label.pack_forget()
            if hasattr(self, 'london_data_entry'):
                self.london_data_entry.pack_forget()
            if hasattr(self, 'london_browse_button'):
                self.london_browse_button.pack_forget()
            # Hide F4 DLC path
            if hasattr(self, 'f4_label'):
                self.f4_label.pack_forget()
            if hasattr(self, 'f4_entry'):
                self.f4_entry.pack_forget()
            if hasattr(self, 'f4_browse_button'):
                self.f4_browse_button.pack_forget()
            # Hide F4VR path
            if hasattr(self, 'f4vr_label'):
                self.f4vr_label.pack_forget()
            if hasattr(self, 'f4vr_entry'):
                self.f4vr_entry.pack_forget()
            if hasattr(self, 'f4vr_browse_button'):
                self.f4vr_browse_button.pack_forget()
            logging.info("Hid London, F4 DLC and F4VR path fields for update mode")
        except Exception as e:
            logging.warning(f"Error hiding DLC and F4VR fields: {e}")
    
    def _show_dlc_and_f4vr_fields(self):
        """Show London and F4VR path fields for fresh install mode (F4 DLC is conditional)"""
        try:
            # Show London path first
            if hasattr(self, 'london_label'):
                self.london_label.pack(pady=(10, 5))
            if hasattr(self, 'london_data_entry'):
                self.london_data_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'london_browse_button'):
                self.london_browse_button.pack()
            # Show F4VR path second
            if hasattr(self, 'f4vr_label'):
                self.f4vr_label.pack(pady=(10, 5))
            if hasattr(self, 'f4vr_entry'):
                self.f4vr_entry.pack(pady=self.get_scaled_value(5))
            if hasattr(self, 'f4vr_browse_button'):
                self.f4vr_browse_button.pack()
            # F4 DLC path is NOT shown here - it's conditional based on DLC detection
            logging.info("Showed London and F4VR path fields for fresh install mode")
        except Exception as e:
            logging.warning(f"Error showing London and F4VR fields: {e}")
    
    def get_player_name_from_save(self, install_path):
        """Extract player name from the latest Fallout 4 VR save file"""
        try:
            saves_path = os.path.join(install_path, "profiles", "default", "saves")
            if not os.path.exists(saves_path):
                logging.warning(f"Saves directory not found at: {saves_path}")
                return None
            
            # Find the latest .fos file
            fos_files = [f for f in os.listdir(saves_path) if f.endswith('.fos')]
            if not fos_files:
                logging.warning("No .fos save files found")
                return None
            
            # Get the most recently modified .fos file
            latest_save = max(
                [os.path.join(saves_path, f) for f in fos_files],
                key=os.path.getmtime
            )
            
            # Read the save file in binary mode and extract the player name
            with open(latest_save, 'rb') as f:
                # Read the first 1024 bytes to find the player name
                data = f.read(1024)
                
                # Look for "FO4_SAVEGAME" header and skip it to find the player name
                # The player name appears after the header
                try:
                    # Convert to string to find the header
                    data_str = data.decode('latin-1', errors='ignore')
                    
                    # Find FO4_SAVEGAME header
                    header_pos = data_str.find('FO4_SAVEGAME')
                    if header_pos == -1:
                        logging.warning("FO4_SAVEGAME header not found in save file")
                        return None
                    
                    # Start searching after the header for the player name
                    search_start = header_pos + len('FO4_SAVEGAME')
                    
                    # Find the next word-like string after the header
                    player_name = None
                    i = search_start
                    while i < len(data):
                        byte = data[i]
                        
                        # Skip non-printable characters and spaces
                        if byte < 32 or byte > 126:
                            i += 1
                            continue
                        
                        # Start collecting potential name characters
                        potential_name = []
                        j = i
                        while j < len(data) and len(potential_name) < 30:
                            byte = data[j]
                            if 32 <= byte <= 126:
                                potential_name.append(chr(byte))
                                j += 1
                            else:
                                break
                        
                        # Check if we have a valid name
                        if potential_name:
                            name_str = ''.join(potential_name).strip()
                            # Validate: starts with letter, only alphanumeric/spaces, reasonable length
                            if (name_str and len(name_str) > 1 and len(name_str) < 25 and 
                                name_str[0].isalpha() and 
                                all(c.isalnum() or c == ' ' for c in name_str)):
                                player_name = name_str
                                break
                        
                        i = j if j > i else i + 1
                    
                    if player_name:
                        logging.info(f"Extracted player name from save: {player_name}")
                        return player_name
                    else:
                        logging.warning("Could not extract player name from save file")
                        return None
                        
                except Exception as e:
                    logging.warning(f"Error parsing save file: {e}")
                    return None
                
        except Exception as e:
            logging.warning(f"Error extracting player name from save: {e}")
            return None

    def _reorganize_ui_for_update(self):
        """Reorganize UI for update mode: move atkins image and installation directory to the top, hide F4/F4VR paths"""
        try:
            # Unpack all widgets from content_frame to reorder them
            if hasattr(self, 'content_frame') and self.content_frame:
                for widget in self.content_frame.winfo_children():
                    widget.pack_forget()
            
            # Hide the original bottom_frame completely - we'll create new widgets in content_frame
            if hasattr(self, 'bottom_frame'):
                self.bottom_frame.pack_forget()
            
            # Explicitly hide the fresh install instruction label and button
            if hasattr(self, 'instruction_label') and self.instruction_label:
                try:
                    self.instruction_label.pack_forget()
                except:
                    pass
            if hasattr(self, 'install_button') and self.install_button:
                try:
                    self.install_button.pack_forget()
                except:
                    pass
            
            bg_color = "#1e1e1e"
            fg_color = "#ffffff"
            
            # Show welcome message above atkins image
            if hasattr(self, 'installation_status_label'):
                self.installation_status_label.pack(pady=(0, self.get_scaled_value(5)), fill="x")
            
            # Show atkins image with no top padding, expandable to fill space
            if hasattr(self, 'atkins_label') and self.atkins_label:
                self.atkins_label.pack(pady=(0, 0), fill="both")
            
            # Show installation directory centered in available space between image and button
            if hasattr(self, 'installation_dir_container'):
                # Update the label text for update mode
                if hasattr(self, 'installation_dir_label'):
                    self.installation_dir_label.config(text="Game Location (Browse for fresh install)")
                self.installation_dir_container.pack(pady=(0, self.get_scaled_value(3)))
            
            # Create new instruction label and button inside content_frame for update mode
            # Use scaled padding for low resolution screens
            # Check if widgets exist and are valid, otherwise create new ones
            create_instruction = True
            create_button = True
            
            if hasattr(self, 'update_instruction_label') and self.update_instruction_label:
                try:
                    if self.update_instruction_label.winfo_exists():
                        create_instruction = False
                except:
                    pass
            
            if hasattr(self, 'update_button') and self.update_button:
                try:
                    if self.update_button.winfo_exists():
                        create_button = False
                except:
                    pass
            
            if create_instruction:
                self.update_instruction_label = tk.Label(self.content_frame, text="Press Update or Enter to continue",
                                                        font=self.regular_font, bg=bg_color, fg="#ffffff")
            self.update_instruction_label.pack(pady=(self.get_scaled_value(10), self.get_scaled_value(5)))
            
            if create_button:
                self.update_button = tk.Button(self.content_frame, text="Update", command=self.validate_and_install,
                                              font=self.regular_font, bg="#28a745", fg=fg_color, bd=0,
                                              relief="flat", activebackground="#218838", padx=self.get_scaled_value(10), pady=self.get_scaled_value(5))
            self.update_button.pack(pady=self.get_scaled_value(3))
            
            # Bind Enter key to trigger the Update button
            self.root.bind('<Return>', lambda event: self.validate_and_install())
            
            logging.info("Reorganized UI for update mode: created new button in content_frame")
        except Exception as e:
            logging.warning(f"Error reorganizing UI for update: {e}")

    def detect_paths(self):
        """Detect Fallout 4, Fallout 4 VR, and existing installation paths"""
        # First, check for existing Fallout London VR installation (unless skip_update_detection is set)
        existing_install = None
        if not self.skip_update_detection:
            existing_install = self.detect_existing_installation()
        else:
            logging.info("Skipping update detection due to --fresh-install flag")
        
        if existing_install:
            self.is_update_detected = True
            self.update_mode = True
            self.detected_install_path = existing_install
            self.mo2_path.set(existing_install)
            
            # Check installed London version and look for 1.03 upgrade opportunity
            installed_version = self.get_installed_london_version(existing_install)
            if installed_version == "1.02":
                logging.info("Installed London version is 1.02, scanning for 1.03 files...")
                london_103_path = self.scan_for_london_103_files()
                if london_103_path:
                    # Prompt user about upgrade - do this on main thread
                    def prompt_upgrade():
                        result = messagebox.askyesno(
                            "Fallout: London 1.03 Available",
                            "Your installation has Fallout: London 1.02, but version 1.03 (Rabbit & Pork) "
                            f"was found at:\n\n{london_103_path}\n\n"
                            "Would you like to upgrade to 1.03 during this update?\n\n"
                            "This will copy the new London files to your installation.\n\n"
                            "Although it is recommended to start a new save when upgrading to 1.03, "
                            "this will not affect your existing saves or mods.",
                            icon='question'
                        )
                        if result:
                            self.upgrade_to_103 = True
                            self.london_103_source_path = london_103_path
                            logging.info(f"User chose to upgrade to London 1.03 from {london_103_path}")
                        else:
                            self.upgrade_to_103 = False
                            logging.info("User declined London 1.03 upgrade")
                    
                    self.root.after(100, prompt_upgrade)
            
            # Extract player name from latest save file
            player_name = self.get_player_name_from_save(existing_install)
            welcome_message = f"Welcome back, {player_name}." if player_name else "Welcome back, Wayfarer."
            
            self.root.after(0, lambda: self.installation_status_label.config(
                text=welcome_message, fg="#ffffff") if self.installation_status_label.winfo_exists() else None)
            self.root.after(0, lambda: self.install_button.config(text="Update") if self.install_button.winfo_exists() else None)
            self.root.after(0, lambda: self.header_label.config(text="Fallout: London VR 0.99") if hasattr(self, 'header_label') and self.header_label.winfo_exists() else None)
            
            # Hide Fallout 4 and Fallout 4 VR path fields for update mode
            self.root.after(0, lambda: self._hide_dlc_and_f4vr_fields())
            
            # Reorganize UI for update mode: move installation directory to bottom and show atkins image
            self.root.after(0, lambda: self._reorganize_ui_for_update())
            
            # Show window after UI is reorganized for update mode
            self.root.after(50, self.root.deiconify)
            
            logging.info(f"Existing installation detected at: {existing_install}")
            # For update mode, skip F4/DLC/London detection and return early
            return
        
        # No existing install, proceed with fresh install detection
        # Get all active drives using the enhanced detection method
        active_drives = self.get_all_active_drives()
        
        # Build comprehensive search paths
        search_paths = []
        
        # First, add registry-detected paths (most reliable)
        steam_paths = self.detect_steam_paths()
        for steam_path in steam_paths:
            # Add the steamapps/common folder
            common_path = os.path.join(steam_path, "steamapps", "common")
            if common_path not in search_paths:
                search_paths.append(common_path)
                logging.info(f"Added Steam path from registry: {common_path}")
        
        gog_paths = self.detect_gog_paths()
        for gog_path in gog_paths:
            if gog_path not in search_paths:
                search_paths.append(gog_path)
                logging.info(f"Added GOG path from registry: {gog_path}")
        
        # Then add common paths on all drives as fallback
        for drive in active_drives:
            # Steam library locations
            search_paths.extend([
                f"{drive}\\Program Files (x86)\\Steam\\steamapps\\common",
                f"{drive}\\Steam\\steamapps\\common",
                f"{drive}\\SteamLibrary\\steamapps\\common",
                f"{drive}\\Games\\Steam\\steamapps\\common"
            ])
            
            # GOG Galaxy locations
            search_paths.extend([
                f"{drive}\\Program Files (x86)\\GOG Galaxy\\Games",
                f"{drive}\\GOG Games",
                f"{drive}\\Games\\GOG"
            ])
            
            # Epic Games Store locations
            search_paths.extend([
                f"{drive}\\Program Files\\Epic Games",
                f"{drive}\\Epic Games"
            ])
            
            # Microsoft Store/Xbox Game Pass locations
            search_paths.extend([
                f"{drive}\\Program Files\\ModifiableWindowsApps",
                f"{drive}\\XboxGames"
            ])
            
            # Other common game locations
            search_paths.extend([
                f"{drive}\\Program Files\\Bethesda.net Launcher\\games",
                f"{drive}\\Program Files (x86)\\Bethesda.net Launcher\\games",
                f"{drive}\\Games",
                f"{drive}\\Program Files\\Games",
                f"{drive}\\Program Files (x86)\\Games"
            ])
        
        # Remove duplicates while preserving order
        unique_paths = []
        for path in search_paths:
            if path not in unique_paths:
                unique_paths.append(path)
        
        logging.info(f"Scanning {len(unique_paths)} paths across {len(active_drives)} drives for Fallout installations")
        
        # Scan all paths for installations
        self.scan_for_fallout_installations(unique_paths)
        
        # Show window after fresh install UI is ready
        self.root.after(10, self.root.deiconify)

    def browse_mo2_path(self):
        """Browse for MO2 installation path with immediate disk space update"""
        path = filedialog.askdirectory(parent=self.root)
        if path:
            try:
                path = self.sanitize_path(path)
                
                # Check if the selected path contains an existing installation
                if self.is_valid_existing_installation(path):
                    # Switch to update mode
                    logging.info(f"User selected existing installation path: {path}")
                    self.is_update_detected = True
                    self.update_mode = True
                    self.detected_install_path = path
                    self.mo2_path.set(path)
                    
                    # Get player name for welcome message
                    player_name = self.get_player_name_from_save(path)
                    welcome_message = f"Welcome back, {player_name}." if player_name else "Welcome back, Wayfarer."
                    
                    # Update UI for update mode
                    if hasattr(self, 'installation_status_label') and self.installation_status_label.winfo_exists():
                        self.installation_status_label.config(text=welcome_message, fg="#ffffff")
                    if hasattr(self, 'install_button') and self.install_button.winfo_exists():
                        self.install_button.config(text="Update")
                    if hasattr(self, 'header_label') and self.header_label.winfo_exists():
                        self.header_label.config(text="Fallout: London VR 0.99")
                    
                    # Hide path fields and reorganize UI for update mode
                    self._hide_dlc_and_f4vr_fields()
                    self._reorganize_ui_for_update()
                    
                    logging.info(f"Switched to update mode for existing installation at: {path}")
                    return
                
                self.mo2_path.set(path)
                # Immediately update disk space display
                self.update_disk_space()
                logging.info(f"MO2 installation path selected: {path}")
            except ValueError as e:
                messagebox.showerror("Invalid Path", str(e))

    def browse_path(self, path_var):
        """Browse for path with enhanced validation and DLC status updates"""
        # Don't allow browsing for London data if waiting for F4 path or if already installed in update mode
        if path_var == self.london_data_path:
            if self.london_data_path.get() == "Waiting for Fallout 4 path":
                return
            # Only block browsing for london_installed in update mode, not fresh install
            if self.london_installed and self.is_update_detected:
                return
        path = filedialog.askdirectory(parent=self.root)
        if path:
            try:
                path = self.sanitize_path(path)
                path_var.set(path)
                if path_var == self.f4_path:
                    # Validate DLC path recursively
                    self.check_dlc_status_independent()
                    # Only check for London in DLC path if user hasn't already selected a valid London path
                    london_path = self.london_data_path.get()
                    if not london_path or london_path in ["", "Waiting for Fallout 4 path", "Already installed"] or not os.path.exists(london_path):
                        self.check_london_installation(path) # Check if London in DLC path
                    if self.missing_dlc:
                        logging.error(f"Invalid Fallout 4 DLC path: missing DLC files: {', '.join(self.missing_dlc)}")
                        if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                            self.root.after(0, lambda: self.dlc_status_label.config(text="Fallout 4: Missing DLC", fg="#ff6666"))
                        if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                            self.root.after(0, lambda: self.message_label.config(text="", fg="#ff6666"))
                    else:
                        logging.info(f"Valid Fallout 4 DLC path selected: {path}")
                        if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                            self.root.after(0, lambda: self.dlc_status_label.config(text="Fallout 4: Ready for installation", fg="#00ff00"))
                    
                    # Check if this path is also a valid Fallout 4 VR path
                    f4vr_exe = os.path.join(path, "Fallout4VR.exe")
                    if os.path.exists(f4vr_exe):
                        self.f4vr_path.set(path)
                        self.root.after(0, lambda: self.f4vr_status_label.config(text="Fallout 4 VR: Ready for installation", fg="#00ff00") if self.f4vr_status_label.winfo_exists() else None)
                        logging.info(f"Updated f4vr_path to {path} as it contains Fallout4VR.exe")
                    else:
                        # Ensure VR path is cleared if it was previously set to this path
                        if self.f4vr_path.get() == path:
                            self.f4vr_path.set("")
                            self.root.after(0, lambda: self.f4vr_status_label.config(text="Fallout 4 VR: Not detected", fg="#ff6666") if self.f4vr_status_label.winfo_exists() else None)
                            logging.info(f"Cleared f4vr_path as {path} no longer contains Fallout4VR.exe")

                elif path_var == self.f4vr_path:
                    f4vr_exe = os.path.join(path, "Fallout4VR.exe")
                    if os.path.exists(f4vr_exe):
                        # Use comprehensive validation for F4VR files
                        is_valid, status_msg, missing_files = self.validate_f4vr_files(path)
                        
                        if is_valid:
                            self.root.after(0, lambda msg=status_msg: self.f4vr_status_label.config(text=msg, fg="#00ff00") if self.f4vr_status_label.winfo_exists() else None)
                            self.root.after(0, lambda: self.message_label.config(text="", fg="#ffffff") if self.message_label.winfo_exists() else None)
                            logging.info("Fallout 4 VR validated successfully")
                            self.check_dlc_status_independent() # Re-check DLC status
                        else:
                            self.root.after(0, lambda msg=status_msg: self.f4vr_status_label.config(text=msg, fg="#ff6666") if self.f4vr_status_label.winfo_exists() else None)
                            if missing_files:
                                missing_str = ", ".join(missing_files[:5])
                                if len(missing_files) > 5:
                                    missing_str += f" and {len(missing_files) - 5} more"
                                self.root.after(0, lambda m=missing_str: self.message_label.config(text=f"Missing: {m}", fg="#ff6666") if self.message_label.winfo_exists() else None)
                            logging.error(f"Fallout 4 VR incomplete: {status_msg}")
                    else:
                        self.root.after(0, lambda: self.f4vr_status_label.config(text="Fallout 4 VR: Invalid location", fg="#ff6666") if self.f4vr_status_label.winfo_exists() else None)
                        self.root.after(0, lambda: self.message_label.config(text="Please select a folder containing Fallout4VR.exe", fg="#ff6666") if self.message_label.winfo_exists() else None)
                        logging.info("Updated message_label for invalid F4VR path")
                elif path_var == self.london_data_path:
                    # Validate Fallout: London path using comprehensive validation
                    if not os.path.exists(path):
                        self.root.after(0, lambda: self.london_status_label.config(text="Fallout: London: Invalid location", fg="#ff6666") if self.london_status_label.winfo_exists() else None)
                        self.root.after(0, lambda: self.message_label.config(text="Fallout: London location does not exist.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                        logging.error(f"Invalid Fallout: London path selected: {path}, does not exist")
                        return
                    
                    # Use comprehensive validation
                    is_valid, version, status_msg, missing_files = self.validate_london_files(path)
                    
                    if is_valid:
                        # Use orange for 1.02, green for 1.03
                        color = "#ffa500" if version == "1.02" else "#00ff00"
                        self.root.after(0, lambda msg=status_msg, c=color: self.london_status_label.config(text=msg, fg=c) if self.london_status_label.winfo_exists() else None)
                        self.root.after(0, lambda: self.message_label.config(text="", fg="#ffffff") if self.message_label.winfo_exists() else None)
                        logging.info(f"Valid Fallout: London {version} detected at {path}")
                        # Check DLC status now that London path is set - this will show/hide F4 DLC field
                        self.check_dlc_status_independent()
                    else:
                        self.london_source_path = None
                        self.london_version = None
                        self.root.after(0, lambda msg=status_msg: self.london_status_label.config(text=msg, fg="#ff6666") if self.london_status_label.winfo_exists() else None)
                        if missing_files:
                            missing_str = ", ".join(missing_files[:5])
                            if len(missing_files) > 5:
                                missing_str += f" and {len(missing_files) - 5} more"
                            self.root.after(0, lambda m=missing_str: self.message_label.config(text=f"Missing: {m}", fg="#ff6666") if self.message_label.winfo_exists() else None)
                        logging.error(f"Invalid Fallout: London path selected: {path}, {status_msg}")
                elif path_var == self.mo2_path:
                    self.update_install_mo2_ui()
               
                # Update disk space after any path change
                self.root.after(100, self.update_disk_space)
                # Reorder status labels (green on top, red below) - use longer delay to ensure all status updates complete
                self.root.after(200, self.reorder_status_labels)
                logging.info(f"Manually selected path: {path}")
            except ValueError as e:
                messagebox.showerror("Invalid Path", str(e))
                logging.error(f"Invalid path selected: {path}, error: {e}")

    def validate_and_proceed(self):
        try:
            if not self.f4_path.get() or not os.path.exists(self.f4_path.get()) or not os.path.exists(os.path.join(self.f4_path.get(), "Fallout4.exe")):
                # Don't update message_label here - let the status labels handle F4 errors
                logging.error("Invalid Fallout 4 path provided or missing Fallout4.exe")
                return

            if not self.f4vr_path.get() or not os.path.exists(self.f4vr_path.get()) or not os.path.exists(os.path.join(self.f4vr_path.get(), "Fallout4VR.exe")):
                self.root.after(0, lambda: self.message_label.config(text="", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error("Invalid Fallout 4 VR path provided or missing Fallout4VR.exe")
                return

            if self.missing_dlc:
                self.root.after(0, lambda: self.message_label.config(text="Required Fallout 4 DLC missing. Please install and try again.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error("Required DLC missing, cannot proceed.")
                return

            # Check London data validation - define london_path at the beginning
            london_path = self.london_data_path.get()
            
            if not self.london_installed:
                if not london_path or london_path == "":
                    self.root.after(0, lambda: self.message_label.config(text="Please provide a valid Fallout: London path.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error("No Fallout: London data path provided.")
                    return
                
                # Check for f4se_loader.exe in the root directory
                f4se_loader_path = os.path.join(london_path, "f4se_loader.exe")
                if not os.path.exists(f4se_loader_path):
                    self.root.after(0, lambda: self.message_label.config(text="", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Invalid Fallout: London path: missing f4se_loader.exe in root: {london_path}")
                    return
                
                # Additional validation - check for Data folder
                london_data_folder = os.path.join(london_path, "Data")
                if not os.path.exists(london_data_folder):
                    self.root.after(0, lambda: self.message_label.config(text="", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Invalid Fallout: London path: missing Data folder: {london_path}")
                    return

            # Sanitize paths (only if they are set)
            if self.f4_path.get():
                self.f4_path.set(self.sanitize_path(self.f4_path.get()))
            self.f4vr_path.set(self.sanitize_path(self.f4vr_path.get()))

            if london_path and london_path != "Already installed" and not self.london_installed:
                self.london_data_path.set(self.sanitize_path(london_path))

            # Remove traces safely before proceeding
            try:
                self.f4_path.trace_remove("write", self.f4_trace_id)
                self.f4vr_path.trace_remove("write", self.f4vr_trace_id)
                self.london_data_path.trace_remove("write", self.london_trace_id)
                self.mo2_path.trace_remove("write", self.mo2_trace_id)
            except Exception as e:
                logging.warning(f"Failed to remove traces: {e}")

        except ValueError as e:
            messagebox.showerror("Validation Error", str(e))
        except Exception as e:
            logging.error(f"Validation error: {e}")
            messagebox.showerror("Error", f"An error occurred during validation: {e}")

    def validate_and_install(self):
        """Validate paths and start installation directly"""
        try:
            # If in update mode, skip path validation and proceed directly to installation
            if self.is_update_detected and self.update_mode:
                logging.info("Update mode detected - skipping Fallout 4 DLC and VR path validation")
                # Verify MO2 directory is set and valid
                if not self.mo2_path.get():
                    self.root.after(0, lambda: self.message_label.config(text="Installation directory not found.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error("MO2 installation directory not set")
                    return
                
                # Check if MO2 path exists and contains valid installation
                mo2_path = self.mo2_path.get()
                mo2_exe = os.path.join(mo2_path, "ModOrganizer.exe")
                if not os.path.exists(mo2_exe):
                    self.root.after(0, lambda: self.message_label.config(text="ModOrganizer.exe not found in installation directory.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"ModOrganizer.exe not found at {mo2_exe}")
                    return
                
                # Set src_london_data to None for update mode (use existing London)
                src_london_data = None
                
                # Remove traces safely before proceeding
                try:
                    if hasattr(self, 'f4_trace_id'):
                        self.f4_path.trace_remove("write", self.f4_trace_id)
                    if hasattr(self, 'f4vr_trace_id'):
                        self.f4vr_path.trace_remove("write", self.f4vr_trace_id)
                    if hasattr(self, 'london_trace_id'):
                        self.london_data_path.trace_remove("write", self.london_trace_id)
                    if hasattr(self, 'mo2_trace_id'):
                        self.mo2_path.trace_remove("write", self.mo2_trace_id)
                except Exception as e:
                    logging.warning(f"Failed to remove traces: {e}")
                
                # Start update process using the dedicated update flow
                self.start_update_process()
                return
            
            # Fresh install mode - validate all paths
            # Validate Fallout 4 VR path
            if not self.f4vr_path.get() or not os.path.exists(self.f4vr_path.get()) or not os.path.exists(os.path.join(self.f4vr_path.get(), "Fallout4VR.exe")):
                self.root.after(0, lambda: self.message_label.config(text="Please provide a valid Fallout 4 VR path.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error("Invalid Fallout 4 VR path provided or missing Fallout4VR.exe")
                return
            
            # Validate Fallout4.esm in f4vr_path/Data
            f4vr_data_dir = os.path.join(self.f4vr_path.get(), "Data")
            esm_path = os.path.join(f4vr_data_dir, "Fallout4.esm")
            if not os.path.exists(esm_path):
                self.root.after(0, lambda: self.message_label.config(text="Fallout4.esm not found in Fallout 4 VR Data folder.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error(f"Fallout4.esm not found in {f4vr_data_dir}")
                return
            
            # Validate DLC - only check if DLC files are missing (they can be in London path, F4 path, or F4VR path)
            if self.missing_dlc:
                self.root.after(0, lambda: self.message_label.config(
                    text=f"Missing required DLC: {', '.join(self.missing_dlc)}", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error(f"Missing DLC files: {', '.join(self.missing_dlc)}")
                return
            
            # Initialize src_london_data to None
            src_london_data = None
            
            # Validate London files path
            london_path = self.london_data_path.get()
            if not self.london_installed:
                if not london_path or london_path == "":
                    self.root.after(0, lambda: self.message_label.config(text="Please provide a valid Fallout: London files path.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error("No Fallout: London files path provided.")
                    return
                
                # Check if london_path exists
                if not os.path.exists(london_path):
                    self.root.after(0, lambda: self.message_label.config(
                        text="Fallout: London path does not exist.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Fallout: London path does not exist: {london_path}")
                    return
                
                # Check for LondonWorldSpace*.esm in root or Data folder
                london_files_found = False
                try:
                    # Check root first
                    for file in os.listdir(london_path):
                        if file.lower().startswith("londonworldspace") and file.lower().endswith(".esm"):
                            london_files_found = True
                            src_london_data = london_path
                            break
                    # Check Data folder if not found in root
                    if not london_files_found:
                        data_path = os.path.join(london_path, "Data")
                        if os.path.exists(data_path):
                            for file in os.listdir(data_path):
                                if file.lower().startswith("londonworldspace") and file.lower().endswith(".esm"):
                                    london_files_found = True
                                    src_london_data = data_path
                                    break
                except Exception as e:
                    self.root.after(0, lambda: self.message_label.config(
                        text=f"Error accessing Fallout: London path: {str(e)}", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Error accessing Fallout: London path {london_path}: {e}")
                    return
                
                if not london_files_found:
                    self.root.after(0, lambda: self.message_label.config(
                        text="No Fallout: London files (e.g., LondonWorldSpace*.esm) found in root or Data folder.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"No Fallout: London files found in {london_path}" + (f" or Data subfolder" if os.path.exists(os.path.join(london_path, "Data")) else ""))
                    return
            else:
                # London is already installed, set the source path to the F4 path
                # Check if London files are in root or Data folder
                f4_path = self.f4_path.get()
                if self._check_london_files_in_path(f4_path):
                    src_london_data = f4_path
                else:
                    f4_data_path = os.path.join(f4_path, "Data")
                    if os.path.exists(f4_data_path) and self._check_london_files_in_path(f4_data_path):
                        src_london_data = f4_data_path
                    else:
                        # This shouldn't happen if london_installed is True, but handle it anyway
                        logging.warning("London marked as installed but files not found in expected locations")
                        src_london_data = f4_path  # Default to F4 path
            
            # MO2 path validation
            if not self.mo2_path.get():
                messagebox.showerror("Error", "Please select an installation directory.")
                logging.error("No MO2 installation directory provided.")
                return
            
            # Check disk space for MO2 drive
            try:
                drive = os.path.splitdrive(self.mo2_path.get())[0]
                if drive:
                    usage = shutil.disk_usage(drive)
                    free_gb = usage.free / (1024 ** 3)
                    if free_gb < 53:
                        messagebox.showerror("Error", f"Insufficient disk space on drive {drive}\ {free_gb:.2f} GB available. At least 53 GB is required.")
                        logging.error(f"Insufficient disk space on {drive}\ {free_gb:.2f} GB available, 53 GB required.")
                        return
            except Exception as e:
                messagebox.showerror("Error", f"Could not check disk space for {drive}: {e}")
                logging.error(f"Could not check disk space for {drive}: {e}")
                return
            
            # Sanitize paths (only if they are set)
            if self.f4_path.get():
                self.f4_path.set(self.sanitize_path(self.f4_path.get()))
            self.f4vr_path.set(self.sanitize_path(self.f4vr_path.get()))
            if london_path and not self.london_installed:
                self.london_data_path.set(self.sanitize_path(london_path))
            
            # Remove traces safely
            try:
                if hasattr(self, 'f4_trace_id'):
                    self.f4_path.trace_remove("write", self.f4_trace_id)
                if hasattr(self, 'f4vr_trace_id'):
                    self.f4vr_path.trace_remove("write", self.f4vr_trace_id)
                if hasattr(self, 'london_trace_id'):
                    self.london_data_path.trace_remove("write", self.london_trace_id)
                if hasattr(self, 'mo2_trace_id'):
                    self.mo2_path.trace_remove("write", self.mo2_trace_id)
            except Exception as e:
                logging.warning(f"Failed to remove traces: {e}")
            
            # Hide input widgets but keep structure
            # Start the installation process with proper UI setup (like update mode)
            self.start_install_process(src_london_data)
        
        except ValueError as e:
            messagebox.showerror("Validation Error", str(e))
            logging.error(f"Validation error: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during validation: {e}")
            logging.error(f"Validation error: {e}")

    def start_install_process(self, src_london_data):
        """Start the installation process for fresh install - similar to start_update_process"""
        # Store the source path for London files (root or Data)
        self.london_source_path = src_london_data
        
        # Hide all widgets
        for widget in self.root.winfo_children():
            if not isinstance(widget, tk.Frame) or widget.winfo_class() != "Frame":
                continue
            for child in widget.winfo_children():
                child.pack_forget()
        
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Add custom title bar
        self.add_custom_title_bar()
        
        bg_color = "#1e1e1e"
        fg_color = "#ffffff"
        
        if self.bg_image:
            bg_label = tk.Label(self.root, image=self.bg_image)
            bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            bg_label.lower()
            main_container = tk.Frame(self.root)
            content_frame = tk.Frame(main_container)
        else:
            self.root.configure(bg=bg_color)
            main_container = tk.Frame(self.root, bg=bg_color)
            content_frame = tk.Frame(main_container, bg=bg_color)
        
        main_container.pack(fill="both", expand=True, padx=20, pady=5)
        content_frame.pack(fill="both", expand=True)
        
        if self.logo:
            tk.Label(content_frame, image=self.logo, bg=bg_color if not self.bg_image else '#1e1e1e').pack(pady=self.get_scaled_value(5))
        
        tk.Label(content_frame, text="Installing Fallout: London VR", font=self.title_font,
                bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color).pack(pady=self.get_scaled_value(5))
        
        self.message_label = tk.Label(content_frame, text="Preparing installation...", font=self.regular_font,
                                      bg=bg_color if not self.bg_image else '#1e1e1e', fg=fg_color,
                                      wraplength=460, justify="center")
        self.message_label.pack(pady=self.get_scaled_value(2))
        
        # Initialize and start slideshow
        self.initialize_and_start_slideshow()
        
        # Start installation in background thread
        threading.Thread(target=self.perform_installation, daemon=True).start()

    def update_disk_space(self, *args):
        """Update disk space display for MO2 installation drive only"""
        try:
            # Don't show disk space for update mode
            if self.is_update_detected and self.update_mode:
                if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                    self.root.after(0, lambda: self.disk_space_label.config(text="", fg="#ffffff"))
                return
            
            mo2_path = self.mo2_path.get()
            if not mo2_path:
                if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                    self.root.after(0, lambda: self.disk_space_label.config(text="", fg="#ffffff"))
                return
            
            drive = os.path.splitdrive(mo2_path)[0]
            if not drive:
                return
            
            usage = shutil.disk_usage(drive)
            free_gb = usage.free / (1024 ** 3)
            
            text = f"Available disk space on drive {drive}\ {free_gb:.2f} GB (Required: 53 GB)"
            color = "#ff6666" if free_gb < 53 else "#ffffff"
            
            if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                self.root.after(0, lambda t=text, c=color: self.disk_space_label.config(text=t, fg=c))
            
            logging.debug(f"Drive {drive}: {free_gb:.2f} GB free")
        except Exception as e:
            logging.error(f"Error updating disk space: {e}")
            if hasattr(self, 'disk_space_label') and self.disk_space_label and self.disk_space_label.winfo_exists():
                self.root.after(0, lambda: self.disk_space_label.config(text="Unable to check disk space.", fg="#ff6666"))

    def check_fallout4_esm_size(self, f4_path):
        """Check if Fallout4.esm needs patching based on its size"""
        # Check both root and Data directories
        for check_path in [f4_path, os.path.join(f4_path, "Data")]:
            esm_path = os.path.join(check_path, "Fallout4.esm")
            if os.path.exists(esm_path):
                esm_size = os.path.getsize(esm_path)
                # These are the sizes that need patching
                if esm_size == 330777465 or esm_size == 330553163:
                    return True, esm_size  # Needs patching
                # Any other size means it doesn't need patching (or unknown)
                return False, esm_size
        return None, 0  # ESM not found

    def check_version(self, f4_path):
        """Check version of Fallout 4 and DLC files with version detection"""
        try:
            data_dir = os.path.join(f4_path, "Data")
            f4_exe_path = os.path.join(f4_path, "Fallout4.exe")
            
            if not os.path.exists(f4_exe_path):
                # Don't update message_label for F4 path errors - handled by status labels
                logging.error("Fallout4.exe not found for version detection.")
                if hasattr(self, 'next_button_message') and self.next_button_message:
                    self.next_button_message.pack_forget()
                return

            # Get Fallout 4 version information
            version, is_next_gen = self.get_fallout4_version(f4_exe_path)
            self.f4_version = version
            
            # Build Fallout 4 status message and color
            f4_status = "Fallout 4: Ready for installation"
            f4_color = "#00ff00"  # Green for ready
            
            required_dlc = {
                "Automatron": "DLCRobot.esm",
                "Wasteland Workshop": "DLCworkshop01.esm", 
                "Far Harbor": "DLCCoast.esm",
                "Contraptions Workshop": "DLCworkshop02.esm",
                "Vault-Tec Workshop": "DLCworkshop03.esm",
                "Nuka-World": "DLCNukaWorld.esm"
            }

            # Use smart DLC detection
            dlc_info = self.detect_dlc_in_both_games()
            
            self.missing_dlc = []
            dlc_with_sources = []
            
            for dlc_name, dlc_data in dlc_info.items():
                if not dlc_data["found_in"]:
                    self.missing_dlc.append(dlc_name)
                    self.dlc_status[dlc_name] = "Not Installed"
                else:
                    source = "F4" if dlc_data["found_in"] == "Fallout 4" else "F4VR"
                    ng_status = "NG" if dlc_data["is_next_gen"] else "Pre-NG"
                    dlc_with_sources.append(f"{dlc_name} ({source}-{ng_status})")
                    self.dlc_status[dlc_name] = "Installed"

            # Check if Fallout4.esm needs patching
            esm_needs_patch, esm_size = self.check_fallout4_esm_size(f4_path)
            
            # Build DLC status message and color based on both DLC presence and ESM status
            if self.missing_dlc:
                # DLC is missing - show F4 DLC path input and red message
                dlc_status = f"DLC: Missing ({', '.join(self.missing_dlc)})"
                dlc_color = "#ff6666"  # Red for missing DLC
                # Show F4 DLC widgets so user can select DLC source
                self.root.after(0, self.show_f4_dlc_widgets)
            elif esm_needs_patch:
                # DLC present but ESM needs patching
                dlc_status = "Fallout 4 DLC files: Ready for installation."
                dlc_color = "#00ff00"  # Green for ready
                logging.info(f"Fallout4.esm needs patching (size: {esm_size:,} bytes)")
                # Hide F4 DLC widgets since DLC was found
                self.root.after(0, self.hide_f4_dlc_widgets)
            else:
                # DLC present and ESM is good (or unknown size which we'll handle during install)
                # Check if any DLC BA2 files need downgrading
                needs_downgrade = any(d["is_next_gen"] for d in dlc_info.values() if d["found_in"])
                
                if needs_downgrade:
                    self.needs_downgrade = True
                else:
                    self.needs_downgrade = False
                dlc_status = "Fallout 4 DLC files: Ready for installation"
                dlc_color = "#00ff00"  # Green for ready
                
                # Add source info to logging
                logging.info(f"DLC sources: {', '.join(dlc_with_sources)}")
                # Hide F4 DLC widgets since DLC was found
                self.root.after(0, self.hide_f4_dlc_widgets)

            # Build London status message and color
            if self.london_installed:
                # Check version to display correct message
                london_path = self.london_data_path.get()
                if london_path and os.path.exists(london_path):
                    is_103 = self._check_london_version_103(london_path)
                    if is_103:
                        london_status = "Fallout: London 1.03 Rabbit & Pork: Ready for installation"
                        london_color = "#00ff00"  # Green for 1.03
                    else:
                        london_status = "Fallout: London 1.02: Ready for installation"
                        london_color = "#ffa500"  # Orange for 1.02
                else:
                    london_status = "Fallout: London: Ready for installation"
                    london_color = "#00ff00"  # Green fallback
            else:
                london_status = "Fallout: London not detected. Select the path to the files you downloaded from GOG."
                london_color = "#ff6666"  # Red for not detected

            # Build additional status message for the general message label
            additional_messages = []
            
            # Add specific next button messages based on status
            if not os.path.exists(f4_exe_path):
                if hasattr(self, 'next_button_message') and self.next_button_message:
                    self.next_button_message.config(text="Please select the correct Fallout 4 VR installation path", fg="#ff6666")
            elif self.missing_dlc:
                if hasattr(self, 'next_button_message') and self.next_button_message:
                    self.next_button_message.config(text="All DLC must be installed to continue", fg="#ff6666")
            elif not self.london_installed:
                if hasattr(self, 'next_button_message') and self.next_button_message:
                    self.next_button_message.config(text="Please select the Fallout London files location", fg="#ff6666")
            else:
                if hasattr(self, 'next_button_message') and self.next_button_message:
                    self.next_button_message.pack_forget()
            
            additional_status = "\n".join(additional_messages) if additional_messages else ""

            # Update all status labels with individual colors
            self.update_status_labels(f4_status, dlc_status, london_status, "#ffffff", f4_color, dlc_color, london_color, additional_status)
            
            logging.info(f"Fallout 4 version: {version} (Next-Gen: {is_next_gen}), DLC summary: {'All DLC Found' if not self.missing_dlc else 'Missing DLC'}, Needs downgrade: {self.needs_downgrade}")
            
        except Exception as e:
            self.update_status_labels(f"Version detection failed: {e}", "", "", "#ff6666")
            logging.error(f"Version detection failed: {e}")
            # Update next button message even on error to clear if necessary
            if hasattr(self, 'next_button_message') and self.next_button_message:
                self.next_button_message.config(text="Version detection failed. Please check the installation path.", fg="#ff6666")

    def update_status_labels(self, f4_status, dlc_status, london_status, general_color="#ffffff", f4_color="#ffffff", dlc_color="#ffffff", london_color="#ffffff", additional_status=""):
        """Update individual status labels with separate colors"""
        try:
            if hasattr(self, 'f4_status_label') and self.f4_status_label and self.f4_status_label.winfo_exists():
                self.root.after(0, lambda: self.f4_status_label.config(text=f4_status, fg=f4_color))
            
            if hasattr(self, 'dlc_status_label') and self.dlc_status_label and self.dlc_status_label.winfo_exists():
                self.root.after(0, lambda: self.dlc_status_label.config(text=dlc_status, fg=dlc_color))
            
            if hasattr(self, 'london_status_label') and self.london_status_label and self.london_status_label.winfo_exists():
                self.root.after(0, lambda: self.london_status_label.config(text=london_status, fg=london_color))
            
            # ONLY update message_label if additional_status is provided AND it's not an F4 error
            if additional_status and "Fallout4.exe" not in additional_status and "missing Fallout4.exe" not in additional_status:
                if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                    self.root.after(0, lambda: self.message_label.config(text=additional_status, fg=general_color))
                    logging.info(f"Updated message_label with: {additional_status}")
            elif not additional_status:
                # Clear message_label if no additional status
                if hasattr(self, 'message_label') and self.message_label and self.message_label.winfo_exists():
                    self.root.after(0, lambda: self.message_label.config(text="", fg="#ffffff"))
                    logging.info("Cleared message_label")
            else:
                logging.info(f"Skipped message_label update for F4 error: {additional_status}")
            
            # Reorder status labels so green messages appear on top
            self.root.after(50, self.reorder_status_labels)
                
        except Exception as e:
            logging.warning(f"Error updating status labels: {e}")

    def hide_install_button(self):
        """Hide the install button during installation"""
        try:
            if hasattr(self, 'install_button') and self.install_button:
                self.install_button.pack_forget()
                logging.info("Install button hidden successfully")
        except Exception as e:
            logging.warning(f"Error hiding install button: {e}")

    def hide_installation_widgets(self):
        """Hide all input widgets during installation"""
        try:
            # Hide all input widgets from the main page
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, (tk.Entry, tk.Button)) and child != self.message_label:
                            child.pack_forget()
            
            # Hide specific widgets
            if hasattr(self, 'f4_status_label'):
                self.f4_status_label.pack_forget()
            if hasattr(self, 'f4vr_status_label'):
                self.f4vr_status_label.pack_forget()
            if hasattr(self, 'dlc_status_label'):
                self.dlc_status_label.pack_forget()
            if hasattr(self, 'london_status_label'):
                self.london_status_label.pack_forget()
            if hasattr(self, 'disk_space_label'):
                self.disk_space_label.pack_forget()
            
            logging.info("Installation widgets hidden successfully")
        except Exception as e:
            logging.warning(f"Error hiding installation widgets: {e}")

    def update_install_mo2_ui(self):
        """Update MO2 installation UI"""
        if hasattr(self, 'message_label') and self.message_label:
            if self.mo2_path.get():
                self.message_label.config(text="Installation directory selected.", fg="#ffffff")

    def perform_downgrade_step(self):
        """Perform the downgrade step automatically with progress bar"""
        try:
            folon_data_dir = os.path.join(self.mo2_path.get(), "mods", "Fallout London Data")
            
            if not os.path.exists(folon_data_dir):
                logging.warning("Fallout London Data directory not found for downgrade")
                return
            
            # Check if any DLC needs downgrading
            downgrader = FalloutVRDowngrader(folon_data_dir)
            dlc_needing_downgrade = downgrader.get_dlc_needing_downgrade()
            
            if not dlc_needing_downgrade:
                logging.info("No DLC needs downgrading, skipping step")
                return
            
            # Count total files to downgrade
            total_files = sum(len(files) for files in dlc_needing_downgrade.values())
            
            self.root.after(0, lambda: self.message_label.config(text=f"Downgrading {total_files} DLC archive files", fg="#ffffff") if self.message_label.winfo_exists() else None)
            self.create_progress_bar("Downgrading DLC Archives")
            
            def progress_update(value):
                self.root.after(0, lambda v=value: self.progress.__setitem__("value", v) if self.progress.winfo_exists() else None)
                self.root.after(0, lambda v=value: self.progress_label.config(text=f"Downgrading DLC Archives ({v:.1f}%)") if self.progress_label.winfo_exists() else None)
            
            downgrader = FalloutVRDowngrader(folon_data_dir, progress_callback=progress_update)
            success_count, downgraded_by_dlc = downgrader.downgrade_dlc_ba2_files()
            
            # Log summary
            if downgraded_by_dlc:
                logging.info(f"Successfully downgraded {success_count} DLC archive files:")
                for dlc_name, files in downgraded_by_dlc.items():
                    logging.info(f"  {dlc_name}: {len(files)} files")
            
            self.root.after(0, lambda: self.message_label.config(text=f"DLC downgrade completed successfully. Downgraded {success_count} files.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("DLC downgrade completed successfully.")
            
            # Small delay to show completion message
            time.sleep(1)
            
        except Exception as e:
            logging.error(f"DLC downgrade failed: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"DLC downgrade failed: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)

    def copy_directory_with_progress(self, src_dir, dest_dir, label_text, exclude_dirs=None):
        """Copy directory with progress tracking and read-only handling"""
        if exclude_dirs is None:
            exclude_dirs = ["source", "scripts\\source"]

        self.create_progress_bar(label_text)
        try:
            src_dir = os.path.normpath(src_dir)
            dest_dir = os.path.normpath(dest_dir)
            os.makedirs(dest_dir, exist_ok=True)

            total_size = 0
            file_count = 0
            files_to_copy = []

            # Calculate total size and build file list
            for root, dirs, files in os.walk(src_dir):
                # Debug: print what we're looking at
                logging.debug(f"Walking directory: {root}")
                logging.debug(f"Subdirectories found: {dirs}")
                
                # Remove F4SE directories (case insensitive)
                original_dirs = dirs[:]
                dirs[:] = [d for d in dirs if d.lower() != 'f4se']
                
                if len(dirs) != len(original_dirs):
                    excluded = [d for d in original_dirs if d not in dirs]
                    logging.info(f"Excluded directories: {excluded} from {root}")
                
                # Skip if we're currently inside an F4SE directory
                if 'f4se' in root.lower():
                    logging.info(f"Skipping F4SE directory: {root}")
                    continue
                
                for file in files:
                    src_file = os.path.join(root, file)
                    rel_path = os.path.relpath(root, src_dir)
                    
                    if rel_path == '.':
                        dest_file = os.path.join(dest_dir, file)
                    else:
                        dest_file = os.path.join(dest_dir, rel_path, file)

                    try:
                        file_size = os.path.getsize(src_file)
                        total_size += file_size
                        file_count += 1
                        files_to_copy.append((src_file, dest_file, file_size))
                    except OSError as e:
                        logging.warning(f"Could not get size of {src_file}: {e}")

            if file_count == 0:
                logging.warning(f"No files found to copy from {src_dir}")
                return

            copied_size = 0
            files_copied = 0

            for src_file, dest_file, file_size in files_to_copy:
                if self.cancel_requested:
                    raise InterruptedError("Installation cancelled by user")

                dest_file_dir = os.path.dirname(dest_file)
                os.makedirs(dest_file_dir, exist_ok=True)

                try:
                    # Handle read-only files in destination
                    if os.path.exists(dest_file):
                        self.remove_readonly_and_overwrite(dest_file)
                    
                    shutil.copy2(src_file, dest_file)
                    copied_size += file_size
                    files_copied += 1

                except PermissionError as e:
                    logging.warning(f"Permission error copying {src_file} to {dest_file}: {e}")
                    # Try to handle read-only file
                    try:
                        self.remove_readonly_and_overwrite(dest_file)
                        shutil.copy2(src_file, dest_file)
                        copied_size += file_size
                        files_copied += 1
                        logging.info(f"Successfully copied after removing read-only: {dest_file}")
                    except Exception as retry_error:
                        logging.error(f"Failed to copy {src_file} after retry: {retry_error}")
                        # Continue with other files instead of failing completely
                        continue

                # Update progress
                progress_percentage = (copied_size / total_size) * 100
                self.root.after(0, lambda pp=progress_percentage: self.progress.__setitem__("value", pp))
                self.root.after(0, lambda pp=progress_percentage: self.progress_label.config(text=f"{label_text} ({pp:.1f}%)"))

            logging.info(f"Successfully copied {files_copied} files from {src_dir} to {dest_dir}")

        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to copy directory: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            logging.error(f"Directory copy failed: {e}")
            raise

    def copy_directory_with_robocopy(self, src_dir, dest_dir, label_text, exclude_dirs=None, exclude_files=None):
        """Parallel copy implementation to replace the hanging robocopy"""
        if exclude_dirs is None:
            exclude_dirs = ["F4SE", "source", "scripts\\source"]
        
        # Use the parallel implementation directly
        return self.copy_directory_parallel(src_dir, dest_dir, label_text, exclude_dirs, exclude_files=exclude_files)

    def copy_directory_parallel(self, src_dir, dest_dir, label_text, exclude_dirs=None, exclude_files=None, max_workers=None):
        """Parallel copy with dynamic worker count based on system resources"""
        if exclude_dirs is None:
            exclude_dirs = ["F4SE", "source", "scripts\\source"]
        if exclude_files is None:
            exclude_files = []
        
        # Dynamic worker calculation
        if max_workers is None:
            cpu_count = os.cpu_count() or 4
            # Consider system load and available memory
            try:
                # Get available memory
                import psutil
                available_memory_gb = psutil.virtual_memory().available / (1024**3)
                
                # Calculate optimal workers
                # Base: 2 workers per CPU core
                base_workers = cpu_count * 2
                
                # Adjust based on available memory (1 worker per 500MB available)
                memory_workers = int(available_memory_gb * 2)
                
                # Take the minimum and cap at 16
                max_workers = min(base_workers, memory_workers, 16)
                max_workers = max(max_workers, 2)  # At least 2 workers
                
                logging.info(f"Dynamic worker count: {max_workers} (CPUs: {cpu_count}, Available RAM: {available_memory_gb:.1f}GB)")
            except:
                # Fallback if psutil not available
                max_workers = min(cpu_count * 2, 8)
                logging.info(f"Fallback worker count: {max_workers} (CPUs: {cpu_count})")
        
        self.create_progress_bar(label_text)
        
        try:
            # Normalize paths
            src_dir = os.path.normpath(src_dir)
            dest_dir = os.path.normpath(dest_dir)
            
            # Collect all files to copy
            files_to_copy = []
            total_size = 0
            exclude_dirs_lower = [d.lower() for d in exclude_dirs]
            exclude_files_lower = [f.lower() for f in exclude_files]
            
            logging.info(f"Scanning directory for parallel copy: {src_dir}")
            if exclude_files:
                logging.info(f"Excluding files: {exclude_files}")
            
            for root, dirs, files in os.walk(src_dir):
                # Remove excluded directories
                dirs[:] = [d for d in dirs if d.lower() not in exclude_dirs_lower]
                
                # Skip if we're in an excluded directory
                if any(excl.lower() in root.lower() for excl in exclude_dirs_lower):
                    continue
                
                for file in files:
                    # Skip excluded files
                    if file.lower() in exclude_files_lower:
                        logging.info(f"Skipping excluded file: {file}")
                        continue
                    
                    src_file = os.path.join(root, file)
                    rel_path = os.path.relpath(src_file, src_dir)
                    dest_file = os.path.join(dest_dir, rel_path)
                    
                    try:
                        size = os.path.getsize(src_file)
                        files_to_copy.append((src_file, dest_file, size))
                        total_size += size
                    except OSError as e:
                        logging.warning(f"Could not get size of {src_file}: {e}")
            
            if not files_to_copy:
                logging.warning(f"No files to copy from {src_dir}")
                return True
            
            logging.info(f"Found {len(files_to_copy)} files to copy, total size: {total_size / (1024**3):.2f} GB")
            
            # Sort by size - copy small files first for better responsiveness
            files_to_copy.sort(key=lambda x: x[2])
            
            # Thread-safe progress tracking
            copied_counter = ThreadSafeCounter(0)
            failed_files = []
            failed_files_lock = threading.Lock()
            files_copied = 0
            
            def copy_single_file_safe(file_info):
                """Copy a single file with thread-safe progress tracking"""
                src_file, dest_file, size = file_info
                
                if self.cancel_requested:
                    raise InterruptedError("Installation cancelled by user")
                
                try:
                    # Create destination directory
                    dest_dir_path = os.path.dirname(dest_file)
                    os.makedirs(dest_dir_path, exist_ok=True)
                    
                    # Handle read-only files
                    if os.path.exists(dest_file):
                        self.remove_readonly_and_overwrite(dest_file)
                    
                    # Copy file in chunks
                    chunk_size = 4 * 1024 * 1024  # 4MB chunks
                    bytes_copied = 0
                    
                    with open(src_file, 'rb') as fsrc:
                        with open(dest_file, 'wb') as fdest:
                            while True:
                                chunk = fsrc.read(chunk_size)
                                if not chunk:
                                    break
                                fdest.write(chunk)
                                bytes_copied += len(chunk)
                                
                                # Thread-safe progress update
                                current_total = copied_counter.add(len(chunk))
                                
                                # Throttle UI updates (only update every 10MB or so)
                                if current_total % (10 * 1024 * 1024) < chunk_size:
                                    progress = (current_total / total_size * 100) if total_size > 0 else 0
                                    if self.progress and self.progress.winfo_exists():
                                        self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                                    if self.progress_label and self.progress_label.winfo_exists():
                                        self.root.after(0, lambda p=progress: self.progress_label.config(
                                            text=f"{label_text} ({p:.1f}%)"
                                        ))
                                
                                # For very large files, periodically flush to disk
                                if bytes_copied % (50 * 1024 * 1024) == 0:  # Every 50MB
                                    fdest.flush()
                                    os.fsync(fdest.fileno())  # Force write to disk
                    
                    # Copy file attributes
                    shutil.copystat(src_file, dest_file)
                    return True, src_file, None
                    
                except Exception as e:
                    logging.error(f"Failed to copy {src_file}: {e}")
                    with failed_files_lock:
                        failed_files.append((src_file, str(e)))
                    return False, src_file, str(e)
            
            # Create thread pool and submit all copy tasks
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit all tasks
                future_to_file = {
                    executor.submit(copy_single_file_safe, file_info): file_info 
                    for file_info in files_to_copy
                }
                
                # Monitor progress
                last_update_time = time.time()
                completed_count = 0
                
                for future in as_completed(future_to_file):
                    completed_count += 1
                    current_time = time.time()
                    
                    try:
                        success, file_path, error = future.result()
                        if success:
                            files_copied += 1
                        # Note: failed_files now handled thread-safely in copy_single_file_safe
                    except Exception as e:
                        file_info = future_to_file[future]
                        with failed_files_lock:
                            failed_files.append((file_info[0], str(e)))
                    
                    # Update progress every 0.5 seconds or every 10 files
                    if current_time - last_update_time > 0.5 or completed_count % 10 == 0:
                        current_copied = copied_counter.get()
                        progress = (current_copied / total_size * 100) if total_size > 0 else 0
                        progress = min(progress, 99.9)  # Cap at 99.9% until fully complete
                        
                        # Calculate speed
                        elapsed = current_time - last_update_time
                        if elapsed > 0:
                            speed_mb = ((current_copied - getattr(self, '_last_copied_size', 0)) / elapsed) / (1024 * 1024)
                            self._last_copied_size = current_copied
                        else:
                            speed_mb = 0
                        
                        # Update UI using direct widget updates
                        if self.progress and self.progress.winfo_exists():
                            self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                        if self.progress_label and self.progress_label.winfo_exists():
                            self.root.after(0, lambda p=progress, s=speed_mb: self.progress_label.config(
                                text=f"{label_text} ({p:.1f}%)"
                            ))
                        
                        last_update_time = current_time
            
            # Final update using direct widget updates
            if self.progress and self.progress.winfo_exists():
                self.root.after(0, lambda: self.progress.__setitem__("value", 100))
            if self.progress_label and self.progress_label.winfo_exists():
                self.root.after(0, lambda: self.progress_label.config(text=f"{label_text} (Complete)"))
            
            # Log results
            logging.info(f"Parallel copy completed: {files_copied}/{len(files_to_copy)} files copied successfully")
            if failed_files:
                logging.warning(f"{len(failed_files)} files failed to copy:")
                for file_path, error in failed_files[:10]:  # Log first 10 failures
                    logging.warning(f"  - {file_path}: {error}")
            
            # If too many failures, raise exception
            if len(failed_files) > len(files_to_copy) * 0.1:  # More than 10% failed
                raise Exception(f"Too many copy failures: {len(failed_files)}/{len(files_to_copy)} files failed")
            
            return True
            
        except Exception as e:
            logging.error(f"Parallel copy failed: {e}")
            if self.message_label and self.message_label.winfo_exists():
                self.root.after(0, lambda es=str(e): self.message_label.config(
                    text=f"Failed to copy directory: {es}", fg="#ff6666"
                ))
            raise

    def get_directory_size(self, path, exclude_dirs=None):
        """Calculate total size of directory for progress tracking"""
        if exclude_dirs is None:
            exclude_dirs = []
        
        total_size = 0
        exclude_dirs_lower = [d.lower() for d in exclude_dirs]
        
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                # Remove excluded directories
                dirnames[:] = [d for d in dirnames if d.lower() not in exclude_dirs_lower]
                
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                    except OSError:
                        pass
            
            logging.info(f"Directory {path} size calculated")
        except Exception as e:
            logging.warning(f"Error calculating directory size: {e}")
            # Return a reasonable estimate if we can't calculate
            return 50 * 1024 * 1024 * 1024  # 50GB estimate
        
        return total_size

    def get_directory_size_and_count(self, path, exclude_dirs=None):
        """Calculate total size and file count of directory for progress tracking"""
        if exclude_dirs is None:
            exclude_dirs = []
        
        total_size = 0
        file_count = 0
        exclude_dirs_lower = [d.lower() for d in exclude_dirs]
        
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                dirnames[:] = [d for d in dirnames if d.lower() not in exclude_dirs_lower]
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                        file_count += 1
                    except OSError:
                        pass
            logging.info(f"Directory {path} analyzed")
        except Exception as e:
            logging.warning(f"Error calculating directory size/count: {e}")
            return 50 * 1024 * 1024 * 1024, 100  # Fallback: 50GB, 100 files
        
        return total_size, file_count

    def copy_mo2_assets(self, update_config=True, exclude_files=None):
        """Copy additional MO2 assets from bundled compressed archive
        
        Args:
            update_config: If True, updates MO2 configuration with paths (fresh install only)
            exclude_files: List of filenames to exclude from copying (e.g., ["ModOrganizer.ini"])
        """
        if exclude_files is None:
            exclude_files = []
        try:
            self.root.after(0, lambda: self.message_label.config(text="Preparing Mod Assets", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            mo2_assets_archive = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "MO2.7z")
            temp_dir = tempfile.mkdtemp()
            
            if os.path.exists(mo2_assets_archive):
                # Create progress bar for extraction
                self.create_progress_bar("Extracting MO2 Assets")
                
                # Get archive size for progress estimation
                try:
                    archive_size = os.path.getsize(mo2_assets_archive)
                except OSError:
                    archive_size = 50 * 1024 * 1024  # Default 50MB estimate
                
                # Estimate extraction time based on archive size (roughly 10MB/sec)
                estimated_time = max(archive_size / (10 * 1024 * 1024), 2.0)  # Minimum 2 seconds
                
                # Use threading for extraction with simulated progress
                extraction_complete = threading.Event()
                extraction_error = None
                
                def extract_thread():
                    nonlocal extraction_error
                    try:
                        logging.info(f"Starting threaded extraction of MO2 assets: {mo2_assets_archive}")
                        with py7zr.SevenZipFile(mo2_assets_archive, mode='r') as z:
                            z.extractall(temp_dir)
                        logging.info(f"Threaded extraction completed")
                    except Exception as e:
                        extraction_error = e
                        logging.error(f"Extraction thread error: {e}")
                    finally:
                        extraction_complete.set()
                
                # Start extraction thread
                extract_thread_obj = threading.Thread(target=extract_thread, daemon=True)
                extract_thread_obj.start()
                
                # Simulate progress while extraction runs
                start_time = time.time()
                while not extraction_complete.is_set():
                    elapsed = time.time() - start_time
                    progress_percent = min((elapsed / estimated_time) * 95, 95)  # Cap at 95% until complete
                    
                    self.root.after(0, lambda p=progress_percent: self.progress.__setitem__("value", p))
                    self.root.after(0, lambda p=progress_percent: self.progress_label.config(text=f"Extracting Fallout: London VR assets ({p:.0f}%)"))
                    
                    time.sleep(0.1)  # Update every 100ms
                
                # Wait for thread to complete
                extract_thread_obj.join(timeout=5.0)  # 5 second timeout
                
                # Check for errors
                if extraction_error:
                    raise extraction_error
                
                # Set to 100% complete
                self.root.after(0, lambda: self.progress.__setitem__("value", 100))
                self.root.after(0, lambda: self.progress_label.config(text="Extracting Fallout: London VR Assets (100%)"))
                
                logging.info(f"Extracted MO2 assets from {mo2_assets_archive} to {temp_dir}")
                
                # Check what was actually extracted and find the correct source directory
                extracted_items = os.listdir(temp_dir)
                logging.info(f"Extracted items: {extracted_items}")
                
                # Look for MO2 directory or use temp_dir directly if files are at root
                mo2_assets_src = os.path.join(temp_dir, "MO2") if "MO2" in extracted_items else temp_dir
                
                # Verify the source directory exists and has content
                if os.path.exists(mo2_assets_src) and os.listdir(mo2_assets_src):
                    # During updates, exclude ModOrganizer.ini to preserve user settings
                    # Use the parameter if provided, otherwise determine based on update_config
                    files_to_exclude = exclude_files if exclude_files else (["ModOrganizer.ini"] if not update_config else None)
                    
                    self.copy_directory_with_robocopy(
                        mo2_assets_src, 
                        self.mo2_path.get(), 
                        "Copying Fallout London VR Files", 
                        exclude_dirs=["source", "scripts\\source"],
                        exclude_files=files_to_exclude
                    )
                    
                    # Only update configuration for fresh installs, not updates
                    if update_config:
                        self.update_mo2_configuration()
                        logging.info("MO2 assets copied and configured successfully")
                    else:
                        logging.info("MO2 assets copied successfully (skipped ModOrganizer.ini and configuration update)")
                else:
                    logging.warning(f"No valid MO2 assets found in extracted archive. Source: {mo2_assets_src}")
                
                # Don't clean up temporary directory - let Windows handle it
                logging.info(f"Temporary directory left for Windows cleanup: {temp_dir}")
            else:
                logging.warning("MO2 assets archive not found; skipping copy.")
                self.root.after(0, lambda: self.message_label.config(text="MO2 assets not found, using defaults.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                
        except Exception as e:
            logging.error(f"Failed to copy MO2 assets: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to copy MO2 assets: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def copy_fallout_data_with_smart_dlc(self):
        """Copy only Fallout London files and DLC files, including Fallout4.esm from f4vr_path if needed"""
        try:
            self.root.after(0, lambda: self.message_label.config(text="Preparing Game Assets", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("Starting smart Fallout data files copy process")
            # Detect DLC in both games
            dlc_info = self.detect_dlc_in_both_games()
            
            # Create mods directory
            mods_dir = os.path.join(self.mo2_path.get(), "mods")
            os.makedirs(mods_dir, exist_ok=True)
            
            # Destination for all files
            london_mod_dir = os.path.join(mods_dir, "Fallout London Data")
            os.makedirs(london_mod_dir, exist_ok=True)
            
            # Step 1: Copy London-specific files
            if self.london_installed:
                # London is in F4 directory, copy only London files
                self.copy_london_files_only(self.f4_path.get(), london_mod_dir)
            else:
                # London is in separate directory, copy from root or Data
                london_path = self.london_data_path.get()
                if hasattr(self, 'london_source_path') and self.london_source_path:
                    src_london_data = self.london_source_path
                    if os.path.exists(src_london_data):
                        # Validate presence of London-specific files
                        london_files_found = False
                        for file in os.listdir(src_london_data):
                            if file.lower().startswith("londonworldspace") and file.lower().endswith(".esm"):
                                london_files_found = True
                                break
                        if not london_files_found:
                            self.root.after(0, lambda: self.message_label.config(
                                text="No Fallout: London files (e.g., LondonWorldSpace*.esm) found in selected path.", fg="#ff6666"
                            ) if self.message_label.winfo_exists() else None)
                            logging.error(f"No Fallout: London files found in {src_london_data}")
                            raise FileNotFoundError(f"No Fallout: London files found in {src_london_data}")
                        # Define files to exclude (Fallout4 BA2 files)
                        exclude_files = [f for f in os.listdir(src_london_data) if f.lower().startswith("fallout4") and f.lower().endswith(".ba2")]
                        self.copy_directory_with_file_exclusions(
                            src_london_data, london_mod_dir,
                            "Copying Fallout: London Data",
                            exclude_files=exclude_files
                        )
                    else:
                        self.root.after(0, lambda: self.message_label.config(
                            text="Fallout: London source path not found.", fg="#ff6666"
                        ) if self.message_label.winfo_exists() else None)
                        logging.error(f"Fallout: London source path not found: {src_london_data}")
                        raise FileNotFoundError(f"Fallout: London source path not found: {src_london_data}")
                else:
                    self.root.after(0, lambda: self.message_label.config(
                        text="Fallout: London source path not set.", fg="#ff6666"
                    ) if self.message_label.winfo_exists() else None)
                    logging.error("Fallout: London source path not set")
                    raise FileNotFoundError("Fallout: London source path not set")
            
            # Step 2: Copy Fallout4.esm from f4_path, london_data_path, or f4vr_path
            esm_copied = False
            expected_esm_sizes = {
                330777600: "323,025 KB (standard)",
                330777465: "323,025 KB (variant)",
                330552064: "322,806 KB (standard)",
                330553163: "322,806 KB (variant)"
            }
            size_tolerance = 1000  # Allow ±1000 bytes for size matching
            for src_dir in [self.f4_path.get(), os.path.join(self.london_data_path.get(), "Data") if self.london_data_path.get() else None, self.london_data_path.get() if self.london_data_path.get() else None, os.path.join(self.f4vr_path.get(), "Data")]:
                if src_dir and os.path.exists(src_dir):
                    src_esm = os.path.join(src_dir, "Fallout4.esm")
                    if os.path.exists(src_esm):
                        esm_size = os.path.getsize(src_esm)
                        # Check if size is within tolerance of expected sizes
                        for expected_size, size_name in expected_esm_sizes.items():
                            if abs(esm_size - expected_size) <= size_tolerance:
                                dest_esm = os.path.join(london_mod_dir, "Fallout4.esm")
                                try:
                                    if os.path.exists(dest_esm):
                                        self.remove_readonly_and_overwrite(dest_esm)
                                    shutil.copy2(src_esm, dest_esm)
                                    logging.info(f"Copied Fallout4.esm from {src_dir} to {london_mod_dir} (size: {esm_size:,} bytes, {esm_size/(1024*1024):.2f} MB)")
                                    esm_copied = True
                                    break
                                except Exception as e:
                                    logging.warning(f"Failed to copy Fallout4.esm from {src_dir}: {e}")
                        if esm_copied:
                            break
            if not esm_copied:
                self.root.after(0, lambda: self.message_label.config(
                    text="No valid Fallout4.esm found in any source path.", fg="#ff6666"
                ) if self.message_label.winfo_exists() else None)
                logging.error(f"No Fallout4.esm with valid size found in f4_path, london_data_path, or f4vr_path/Data. Expected sizes: {', '.join([f'{size:,} bytes ({name})' for size, name in expected_esm_sizes.items()])} ±{size_tolerance} bytes")
                raise FileNotFoundError("No valid Fallout4.esm found")
            
            # Step 3: Copy DLC files from best sources
            self.copy_best_dlc_files(dlc_info, london_mod_dir)
            
            # Step 4: Apply ESM patch if needed
            patcher = ESMPatcher(self)
            patch_result = patcher.patch_fallout4_esm(london_mod_dir)
            
            if not patch_result:
                # Patching failed but don't raise an exception
                # The game might still work with the unpatched ESM
                logging.warning("ESM patching failed, continuing with unpatched ESM")
            
            self.root.after(0, lambda: self.message_label.config(text="Game files copied successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("Smart Fallout data files copy completed")
        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(
                text=f"Failed to copy game data: {es}", fg="#ff6666"
            ) if self.message_label.winfo_exists() else None)
            logging.error(f"Game data copy failed: {e}")
            raise

    def copy_london_files_only(self, f4_path, dest_dir):
        """Copy only London-specific files from Fallout 4 directory"""
        try:
            src_data_dir = os.path.join(f4_path, "Data")
            
            # Define London-specific file patterns
            london_patterns = [
                "London*.esm",
                "London*.esp",
                "London*.ba2",
                "LondonWorldSpace*",
                "FOLON*"
            ]
            
            # Define directories to exclude
            exclude_dirs = ["f4se", "F4SE"]
            
            # Build list of files to copy
            files_to_copy = []
            total_size = 0
            
            # Copy files matching London patterns from Data root
            for pattern in london_patterns:
                for file_path in Path(src_data_dir).glob(pattern):
                    try:
                        size = file_path.stat().st_size
                        files_to_copy.append((str(file_path), file_path.name, size))
                        total_size += size
                    except OSError as e:
                        logging.warning(f"Could not get size of {file_path}: {e}")
            
            # Copy all files from Video subdirectory (excluding F4SE)
            video_dir = os.path.join(src_data_dir, "Video")
            if os.path.exists(video_dir):
                for root, dirs, files in os.walk(video_dir):
                    # Exclude F4SE directories
                    dirs[:] = [d for d in dirs if d.lower() not in [e.lower() for e in exclude_dirs]]
                    for file in files:
                        src_file = os.path.join(root, file)
                        rel_path = os.path.relpath(src_file, src_data_dir)
                        try:
                            size = os.path.getsize(src_file)
                            files_to_copy.append((src_file, rel_path, size))
                            total_size += size
                        except OSError as e:
                            logging.warning(f"Could not get size of {src_file}: {e}")
            
            # Copy all files from Scripts subdirectory (excluding F4SE)
            scripts_dir = os.path.join(src_data_dir, "Scripts")
            if os.path.exists(scripts_dir):
                for root, dirs, files in os.walk(scripts_dir):
                    # Exclude F4SE directories
                    dirs[:] = [d for d in dirs if d.lower() not in [e.lower() for e in exclude_dirs]]
                    for file in files:
                        src_file = os.path.join(root, file)
                        rel_path = os.path.relpath(src_file, src_data_dir)
                        try:
                            size = os.path.getsize(src_file)
                            files_to_copy.append((src_file, rel_path, size))
                            total_size += size
                        except OSError as e:
                            logging.warning(f"Could not get size of {src_file}: {e}")
            
            # Copy all files from Textures subdirectory (excluding F4SE)
            textures_dir = os.path.join(src_data_dir, "Textures")
            if os.path.exists(textures_dir):
                for root, dirs, files in os.walk(textures_dir):
                    # Exclude F4SE directories
                    dirs[:] = [d for d in dirs if d.lower() not in [e.lower() for e in exclude_dirs]]
                    for file in files:
                        src_file = os.path.join(root, file)
                        rel_path = os.path.relpath(src_file, src_data_dir)
                        try:
                            size = os.path.getsize(src_file)
                            files_to_copy.append((src_file, rel_path, size))
                            total_size += size
                        except OSError as e:
                            logging.warning(f"Could not get size of {src_file}: {e}")
            
            logging.info(f"Found {len(files_to_copy)} London-specific files to copy, total size: {total_size / (1024**3):.2f} GB")
            
            # Copy the files with progress
            self.create_progress_bar("Copying Fallout: London Data")
            copied_size = 0
            
            for src_file, rel_path, size in files_to_copy:
                if self.cancel_requested:
                    raise InterruptedError("Installation cancelled by user")
                
                # Determine destination path
                if os.path.sep in rel_path:
                    dest_file = os.path.join(dest_dir, rel_path)
                    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                else:
                    dest_file = os.path.join(dest_dir, rel_path)
                
                # Copy file
                try:
                    if os.path.exists(dest_file):
                        self.remove_readonly_and_overwrite(dest_file)
                    shutil.copy2(src_file, dest_file)
                    copied_size += size
                    
                    # Update progress
                    progress = (copied_size / total_size * 100) if total_size > 0 else 0
                    self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                    self.root.after(0, lambda p=progress: self.progress_label.config(
                        text=f"Copying Fallout: London Data ({p:.1f}%)"
                    ))
                except Exception as e:
                    logging.error(f"Failed to copy {src_file}: {e}")
            
        except Exception as e:
            logging.error(f"Failed to copy London files: {e}")
            raise

    def detect_dlc_in_both_games(self):
        """Detect DLC in Fallout 4 and Fallout 4 VR - only check root and Data folders"""
        dlc_info = {
            "Automatron": {
                "esm": "DLCRobot.esm",
                "ba2_prefixes": ["DLCRobot"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            },
            "Wasteland Workshop": {
                "esm": "DLCworkshop01.esm",
                "ba2_prefixes": ["DLCworkshop01"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            },
            "Far Harbor": {
                "esm": "DLCCoast.esm",
                "ba2_prefixes": ["DLCCoast"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            },
            "Contraptions Workshop": {
                "esm": "DLCworkshop02.esm",
                "ba2_prefixes": ["DLCworkshop02"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            },
            "Vault-Tec Workshop": {
                "esm": "DLCworkshop03.esm",
                "ba2_prefixes": ["DLCworkshop03"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            },
            "Nuka-World": {
                "esm": "DLCNukaWorld.esm",
                "ba2_prefixes": ["DLCNukaWorld"],
                "found_in": None,
                "esm_path": None,
                "ba2_paths": [],
                "is_next_gen": None
            }
        }
        
        # Check Fallout 4 DLC path (check root and Data subfolder only)
        f4_dlc_dir = Path(self.f4_path.get())
        if f4_dlc_dir.exists():
            # Check root directory
            self._scan_single_directory_for_dlc(f4_dlc_dir, dlc_info, "Fallout 4 DLC")
            
            # Check Data subdirectory if it exists
            f4_data_dir = f4_dlc_dir / "Data"
            if f4_data_dir.exists():
                self._scan_single_directory_for_dlc(f4_data_dir, dlc_info, "Fallout 4 DLC")
        
        # Check Fallout 4 VR Data folder only
        if self.f4vr_path.get():
            f4vr_data_dir = Path(self.f4vr_path.get()) / "Data"
            if f4vr_data_dir.exists():
                self._scan_single_directory_for_dlc(f4vr_data_dir, dlc_info, "Fallout 4 VR")
        
        # Check Fallout London path (check root and Data subfolder)
        if self.london_data_path.get() and self.london_data_path.get() not in ["", "Waiting for Fallout 4 path", "Already installed"]:
            london_dir = Path(self.london_data_path.get())
            if london_dir.exists():
                # Check root directory
                self._scan_single_directory_for_dlc(london_dir, dlc_info, "Fallout London")
                
                # Check Data subdirectory if it exists
                london_data_dir = london_dir / "Data"
                if london_data_dir.exists():
                    self._scan_single_directory_for_dlc(london_data_dir, dlc_info, "Fallout London")
        
        return dlc_info

    def _scan_single_directory_for_dlc(self, directory, dlc_info, source_name):
        """Scan a single directory (non-recursive) for DLC files"""
        try:
            # Get all files in this directory only (not subdirectories)
            files = [f for f in directory.iterdir() if f.is_file()]
            
            # Build maps for quick lookup
            esm_map = {}
            ba2_map = {}
            
            for file_path in files:
                file_name = file_path.name
                file_lower = file_name.lower()
                
                if file_lower.endswith('.esm'):
                    esm_map[file_name] = file_path
                elif file_lower.endswith('.ba2'):
                    # Extract prefix
                    prefix = file_name.split(' - ')[0] if ' - ' in file_name else file_name.split('.')[0]
                    if prefix not in ba2_map:
                        ba2_map[prefix] = []
                    ba2_map[prefix].append(file_path)
            
            # Match DLC files
            for dlc_name, dlc_data in dlc_info.items():
                if dlc_data["esm"] in esm_map:
                    esm_path = esm_map[dlc_data["esm"]]
                    is_ng = False
                    ba2_paths = []
                    
                    # Find matching BA2 files
                    for prefix in dlc_data["ba2_prefixes"]:
                        if prefix in ba2_map:
                            for ba2_path in ba2_map[prefix]:
                                ba2_paths.append(ba2_path)
                                # Check if Next-Gen
                                header = BA2Header.from_file(ba2_path)
                                if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                                    is_ng = True
                    
                    # Only update if not found yet or if this is a better version (pre-NG preferred)
                    if dlc_data["found_in"] is None or (not is_ng and dlc_data.get("is_next_gen")):
                        dlc_data["found_in"] = source_name
                        dlc_data["esm_path"] = esm_path
                        dlc_data["ba2_paths"] = ba2_paths
                        dlc_data["is_next_gen"] = is_ng
                        logging.info(f"{dlc_name} found in {source_name} at {directory} (Next-Gen: {is_ng})")
                        
        except Exception as e:
            logging.error(f"Error scanning directory {directory}: {e}")

    def check_dlc_ba2_version(self, data_dir, ba2_prefixes):
        """Check if DLC BA2 files are Next-Gen version"""
        try:
            for prefix in ba2_prefixes:
                ba2_files = list(Path(data_dir).glob(f"{prefix}*.ba2"))
                for ba2_file in ba2_files:
                    header = BA2Header.from_file(ba2_file)
                    if header and header.version in [ArchiveVersionEnum.FALLOUT_4_NG, ArchiveVersionEnum.FALLOUT_4_NG2]:
                        return True
            return False
        except Exception as e:
            logging.warning(f"Error checking BA2 version: {e}")
            return False

    def check_dlc_status_independent(self):
        """Check DLC status in provided paths - only root and Data folders"""
        try:
            data_dirs_to_check = []
            dlc_found_in_london_or_vr = False  # Track if DLC found in London or VR paths
            
            # Fallout 4 DLC path - check if it has a value (auto-detected or user-selected)
            # We need to check this path regardless of widget visibility since it may be auto-detected
            f4_dirs = []
            if self.f4_path.get() and os.path.exists(self.f4_path.get()):
                f4_dirs.append(("Fallout 4 DLC", self.f4_path.get()))
                f4_data = os.path.join(self.f4_path.get(), "Data")
                if os.path.exists(f4_data):
                    f4_dirs.append(("Fallout 4 DLC", f4_data))
            
            # Fallout 4 VR Data folder only
            vr_dirs = []
            if self.f4vr_path.get() and os.path.exists(self.f4vr_path.get()):
                f4vr_data = os.path.join(self.f4vr_path.get(), "Data")
                if os.path.exists(f4vr_data):
                    vr_dirs.append(("Fallout 4 VR", f4vr_data))
            
            # Fallout London path - check root and Data
            london_dirs = []
            london_path_value = self.london_data_path.get()
            if london_path_value and london_path_value not in ["", "Waiting for Fallout 4 path", "Already installed"] and os.path.exists(london_path_value):
                london_dirs.append(("Fallout London", london_path_value))
                london_data = os.path.join(london_path_value, "Data")
                if os.path.exists(london_data):
                    london_dirs.append(("Fallout London", london_data))
            
            # Combine all dirs - check London and VR first, then F4
            data_dirs_to_check = london_dirs + vr_dirs + f4_dirs
            
            if not data_dirs_to_check:
                self.root.after(0, lambda: self.dlc_status_label.config(
                    text="DLC: No paths found to check", fg="#ff6666") if self.dlc_status_label.winfo_exists() else None)
                self.root.after(0, lambda: self.message_label.config(
                    text="No valid paths found to check for DLC.", fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error("No valid paths found to check for DLC")
                # Set all required DLC as missing when no paths to check
                self.missing_dlc = ["Automatron", "Wasteland Workshop", "Far Harbor", "Contraptions Workshop", "Vault-Tec Workshop", "Nuka-World"]
                return
            
            required_dlc = {
                "Automatron": "DLCRobot.esm",
                "Wasteland Workshop": "DLCworkshop01.esm",
                "Far Harbor": "DLCCoast.esm",
                "Contraptions Workshop": "DLCworkshop02.esm",
                "Vault-Tec Workshop": "DLCworkshop03.esm",
                "Nuka-World": "DLCNukaWorld.esm"
            }
            
            found_dlc = {}
            missing_dlc = []
            
            # Check each directory (non-recursive)
            for game_name, dir_path in data_dirs_to_check:
                try:
                    files = {f.lower(): f for f in os.listdir(dir_path) if os.path.isfile(os.path.join(dir_path, f)) and f.lower().endswith('.esm')}
                    for dlc_name, esm_file in required_dlc.items():
                        if esm_file.lower() in files:
                            if dlc_name not in found_dlc:
                                found_dlc[dlc_name] = game_name
                                logging.info(f"Found {dlc_name} ({esm_file}) in {dir_path}")
                                # Track if found in London or VR (not F4 DLC path)
                                if game_name in ["Fallout London", "Fallout 4 VR"]:
                                    dlc_found_in_london_or_vr = True
                except Exception as e:
                    logging.error(f"Error checking directory {dir_path}: {e}")
                    self.root.after(0, lambda: self.message_label.config(
                        text=f"Error accessing DLC path {dir_path}: {str(e)}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            
            # Determine missing DLC
            for dlc_name in required_dlc:
                if dlc_name not in found_dlc:
                    missing_dlc.append(dlc_name)
                    logging.warning(f"DLC missing: {dlc_name}")
            
            # Determine if we need to show F4 DLC widgets
            # Show if: DLC is missing AND not found in London path
            london_valid = london_path_value and london_path_value not in ["", "Waiting for Fallout 4 path", "Already installed"] and os.path.exists(london_path_value)
            vr_valid = self.f4vr_path.get() and os.path.exists(self.f4vr_path.get()) and os.path.exists(os.path.join(self.f4vr_path.get(), "Fallout4VR.exe"))
            
            # Check if ALL DLC was found in London path specifically
            all_dlc_in_london = all(found_dlc.get(dlc) == "Fallout London" for dlc in required_dlc if dlc in found_dlc) and len(found_dlc) == len(required_dlc)
            # Check if ALL DLC was found in London or VR paths
            all_dlc_in_london_or_vr = all(found_dlc.get(dlc) in ["Fallout London", "Fallout 4 VR"] for dlc in required_dlc if dlc in found_dlc) and len(found_dlc) == len(required_dlc)
            
            if missing_dlc:
                self.missing_dlc = missing_dlc
                dlc_status = f"DLC: Missing ({', '.join(missing_dlc)})"
                dlc_color = "#ff6666"
                self.root.after(0, lambda: self.message_label.config(
                    text=f"", fg="#ff6666") if self.message_label.winfo_exists() else None)
                # Show F4 DLC widgets only if London path doesn't have all DLC
                if london_valid and vr_valid and not all_dlc_in_london:
                    self.root.after(0, self.show_f4_dlc_widgets)
                elif all_dlc_in_london:
                    # London has some DLC but not all - still hide F4 DLC since user should fix London path
                    self.root.after(0, self.hide_f4_dlc_widgets)
            else:
                self.missing_dlc = []
                
                # Check if Fallout4.esm needs patching
                esm_needs_patch = False
                if self.f4_path.get():
                    esm_needs_patch, esm_size = self.check_fallout4_esm_size(self.f4_path.get())
                
                if esm_needs_patch:
                    logging.info(f"Fallout4.esm needs patching (size: {esm_size:,} bytes)")
                else:
                    dlc_info = self.detect_dlc_in_both_games()
                    needs_downgrade = any(d["is_next_gen"] for d in dlc_info.values() if d["found_in"])
                    if needs_downgrade:
                        self.needs_downgrade = True
                    else:
                        self.needs_downgrade = False
                
                dlc_status = "Fallout 4 DLC files: Ready for installation"
                dlc_color = "#00ff00"  # Green for ready
                
                # Hide F4 DLC widgets if all DLC was found in London path
                # Show F4 DLC widgets only if DLC was found in F4 path (not London)
                if all_dlc_in_london:
                    # All DLC found in London path - hide F4 DLC widgets
                    self.root.after(0, self.hide_f4_dlc_widgets)
                elif all_dlc_in_london_or_vr:
                    # All DLC found in London or VR paths - hide F4 DLC widgets
                    self.root.after(0, self.hide_f4_dlc_widgets)
                elif self.f4_path.get() and london_valid and vr_valid:
                    # DLC found in F4 path - show widgets with the auto-detected path
                    self.root.after(0, self.show_f4_dlc_widgets)
                else:
                    # Default: hide F4 DLC widgets
                    self.root.after(0, self.hide_f4_dlc_widgets)
            
            self.root.after(0, lambda: self.dlc_status_label.config(text=dlc_status, fg=dlc_color) if self.dlc_status_label.winfo_exists() else None)
            logging.info(f"DLC status: {dlc_status}")
            
            # Reorder status labels so green messages appear on top - use longer delay to ensure label updates complete
            self.root.after(100, self.reorder_status_labels)
        
        except Exception as e:
            logging.error(f"Error checking DLC status: {e}")
            self.root.after(0, lambda: self.dlc_status_label.config(
                text="DLC: Error checking status", fg="#ff6666") if self.dlc_status_label.winfo_exists() else None)
            self.root.after(0, lambda: self.message_label.config(
                text="Error checking DLC status. Please verify paths.", fg="#ff6666") if self.message_label.winfo_exists() else None)

    def get_dlc_files_to_copy(self, dlc_name, esm_file, ba2_prefixes, source_dir):
        """Get list of DLC files to copy (ESM + BA2 files)"""
        files_to_copy = []
        
        # Add ESM file
        esm_path = os.path.join(source_dir, esm_file)
        if os.path.exists(esm_path):
            files_to_copy.append(esm_file)
        
        # Add all BA2 files with matching prefixes
        for prefix in ba2_prefixes:
            ba2_files = list(Path(source_dir).glob(f"{prefix}*.ba2"))
            for ba2_file in ba2_files:
                files_to_copy.append(ba2_file.name)
        
        return files_to_copy

    def should_copy_dlc(self, dlc_name, dlc_data, f4vr_data_dir):
        """Check if we should copy DLC files"""
        # Check if DLC already exists in F4VR
        f4vr_esm = os.path.join(f4vr_data_dir, dlc_data["esm"])
        if os.path.exists(f4vr_esm):
            # Check if F4VR version is already pre-NG
            f4vr_is_ng = self.check_dlc_ba2_version(f4vr_data_dir, dlc_data["ba2_prefixes"])
            if not f4vr_is_ng:
                logging.info(f"{dlc_name} already exists in F4VR as pre-NG version, skipping copy")
                return False
        return True

    def copy_best_dlc_files(self, dlc_info, dest_dir):
        """Copy DLC files from their found locations to dest_dir"""
        self.create_progress_bar("Copying DLC files")
        
        total_dlc = sum(1 for d in dlc_info.values() if d["found_in"])
        processed_dlc = 0
        
        for dlc_name, dlc_data in dlc_info.items():
            if dlc_data["found_in"]:
                files_to_copy = []
                if dlc_data["esm_path"]:
                    files_to_copy.append(dlc_data["esm_path"])
                files_to_copy.extend(dlc_data["ba2_paths"])
                
                logging.info(f"Copying {dlc_name} from {dlc_data['found_in']} (Pre-NG: {not dlc_data['is_next_gen']})")
                
                for file_path in files_to_copy:
                    dest_file = os.path.join(dest_dir, file_path.name)  # Copy to root of dest_dir
                    try:
                        shutil.copy2(str(file_path), dest_file)
                        logging.debug(f"Copied {file_path.name} from {dlc_data['found_in']}")
                    except Exception as e:
                        logging.error(f"Failed to copy {file_path}: {e}")
                
                processed_dlc += 1
                progress = (processed_dlc / total_dlc) * 100
                self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                self.root.after(0, lambda p=progress, n=dlc_name: self.progress_label.config(text=f"Copying DLC: {n} ({p:.0f}%)"))

    def copy_directory_with_file_exclusions(self, src_dir, dest_dir, label_text, exclude_dirs=None, exclude_files=None):
        """Copy directory with file exclusions using parallel method"""
        if exclude_dirs is None:
            exclude_dirs = ["F4SE", "source", "scripts\\source"]
        if exclude_files is None:
            exclude_files = []
        
        # Convert to parallel copy with exclusions
        return self.copy_directory_parallel_with_exclusions(src_dir, dest_dir, label_text, exclude_dirs, exclude_files)

    def copy_directory_parallel_with_exclusions(self, src_dir, dest_dir, label_text, exclude_dirs=None, exclude_files=None, max_workers=8):
        """Parallel copy with file exclusions"""
        if exclude_dirs is None:
            exclude_dirs = ["F4SE", "source", "scripts\\source"]
        if exclude_files is None:
            exclude_files = []
        
        self.create_progress_bar(label_text)
        
        try:
            # Normalize paths
            src_dir = os.path.normpath(src_dir)
            dest_dir = os.path.normpath(dest_dir)
            
            # Collect all files to copy
            files_to_copy = []
            total_size = 0
            exclude_dirs_lower = [d.lower() for d in exclude_dirs]
            exclude_files_set = set(exclude_files)  # Convert to set for faster lookup
            
            logging.info(f"Scanning directory for parallel copy with exclusions: {src_dir}")
            logging.info(f"Excluding files: {exclude_files[:10]}..." if len(exclude_files) > 10 else f"Excluding files: {exclude_files}")
            
            for root, dirs, files in os.walk(src_dir):
                # Remove excluded directories
                dirs[:] = [d for d in dirs if d.lower() not in exclude_dirs_lower]
                
                # Skip if we're in an excluded directory
                if any(excl.lower() in root.lower() for excl in exclude_dirs_lower):
                    continue
                
                for file in files:
                    # Skip excluded files
                    if file in exclude_files_set:
                        continue
                        
                    src_file = os.path.join(root, file)
                    rel_path = os.path.relpath(src_file, src_dir)
                    dest_file = os.path.join(dest_dir, rel_path)
                    
                    try:
                        size = os.path.getsize(src_file)
                        files_to_copy.append((src_file, dest_file, size))
                        total_size += size
                    except OSError as e:
                        logging.warning(f"Could not get size of {src_file}: {e}")
            
            if not files_to_copy:
                logging.warning(f"No files to copy from {src_dir}")
                return True
            
            logging.info(f"Found {len(files_to_copy)} files to copy (after exclusions), total size: {total_size / (1024**3):.2f} GB")
            
            # Sort by size - copy small files first for better responsiveness
            files_to_copy.sort(key=lambda x: x[2])
            
            # Thread-safe progress tracking
            copied_counter = ThreadSafeCounter(0)
            failed_files = []
            failed_files_lock = threading.Lock()
            files_copied = 0
            
            def copy_single_file_safe_exclusions(file_info):
                """Copy a single file with thread-safe progress tracking for exclusions method"""
                src_file, dest_file, size = file_info
                
                if self.cancel_requested:
                    raise InterruptedError("Installation cancelled by user")
                
                try:
                    # Create destination directory
                    dest_dir_path = os.path.dirname(dest_file)
                    os.makedirs(dest_dir_path, exist_ok=True)
                    
                    # Handle read-only files
                    if os.path.exists(dest_file):
                        self.remove_readonly_and_overwrite(dest_file)
                    
                    # Copy file in chunks
                    chunk_size = 4 * 1024 * 1024  # 4MB chunks
                    bytes_copied = 0
                    
                    with open(src_file, 'rb') as fsrc:
                        with open(dest_file, 'wb') as fdest:
                            while True:
                                chunk = fsrc.read(chunk_size)
                                if not chunk:
                                    break
                                fdest.write(chunk)
                                bytes_copied += len(chunk)
                                
                                # Thread-safe progress update
                                current_total = copied_counter.add(len(chunk))
                                
                                # Throttle UI updates (only update every 10MB or so)
                                if current_total % (10 * 1024 * 1024) < chunk_size:
                                    progress = (current_total / total_size * 100) if total_size > 0 else 0
                                    if self.progress and self.progress.winfo_exists():
                                        self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                                    if self.progress_label and self.progress_label.winfo_exists():
                                        self.root.after(0, lambda p=progress: self.progress_label.config(
                                            text=f"{label_text} ({p:.1f}%)"
                                        ))
                                
                                # For very large files, periodically flush to disk
                                if bytes_copied % (50 * 1024 * 1024) == 0:  # Every 50MB
                                    fdest.flush()
                                    os.fsync(fdest.fileno())  # Force write to disk
                    
                    # Copy file attributes
                    shutil.copystat(src_file, dest_file)
                    return True, src_file, None
                    
                except Exception as e:
                    logging.error(f"Failed to copy {src_file}: {e}")
                    with failed_files_lock:
                        failed_files.append((src_file, str(e)))
                    return False, src_file, str(e)
            
            # Create thread pool and submit all copy tasks
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit all tasks
                future_to_file = {
                    executor.submit(copy_single_file_safe_exclusions, file_info): file_info 
                    for file_info in files_to_copy
                }
                
                # Monitor progress
                last_update_time = time.time()
                completed_count = 0
                
                for future in as_completed(future_to_file):
                    completed_count += 1
                    current_time = time.time()
                    
                    try:
                        success, file_path, error = future.result()
                        if success:
                            files_copied += 1
                        else:
                            failed_files.append((file_path, error))
                    except Exception as e:
                        file_info = future_to_file[future]
                        with failed_files_lock:
                            failed_files.append((file_info[0], str(e)))
                    
                    # Update progress every 0.5 seconds or every 10 files
                    if current_time - last_update_time > 0.5 or completed_count % 10 == 0:
                        current_copied = copied_counter.get()
                        progress = (current_copied / total_size * 100) if total_size > 0 else 0
                        progress = min(progress, 99.9)  # Cap at 99.9% until fully complete
                        
                        # Update UI (no speed display for file exclusions to keep it simple)
                        if self.progress and self.progress.winfo_exists():
                            self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                        if self.progress_label and self.progress_label.winfo_exists():
                            self.root.after(0, lambda p=progress: self.progress_label.config(
                                text=f"{label_text} ({p:.1f}%)"
                            ))
                        
                        last_update_time = current_time
            
            # Final update using direct widget updates
            if self.progress and self.progress.winfo_exists():
                self.root.after(0, lambda: self.progress.__setitem__("value", 100))
            if self.progress_label and self.progress_label.winfo_exists():
                self.root.after(0, lambda: self.progress_label.config(text=f"{label_text} (Complete)"))
            
            # Log results
            logging.info(f"Parallel copy with exclusions completed: {files_copied}/{len(files_to_copy)} files copied successfully")
            if failed_files:
                logging.warning(f"{len(failed_files)} files failed to copy:")
                for file_path, error in failed_files[:10]:  # Log first 10 failures
                    logging.warning(f"  - {file_path}: {error}")
            
            # If too many failures, raise exception
            if len(failed_files) > len(files_to_copy) * 0.1:  # More than 10% failed
                raise Exception(f"Too many copy failures: {len(failed_files)}/{len(files_to_copy)} files failed")
            
            return True
            
        except Exception as e:
            logging.error(f"Parallel copy with exclusions failed: {e}")
            if self.message_label and self.message_label.winfo_exists():
                self.root.after(0, lambda es=str(e): self.message_label.config(
                    text=f"Failed to copy directory: {es}", fg="#ff6666"
                ))
            raise

    def update_mo2_configuration(self):
        """Update MO2 configuration files with correct paths"""
        try:
            mo2_ini_path = os.path.join(self.mo2_path.get(), "ModOrganizer.ini")
           
            # Get paths from user selections
            f4vr_path = self.f4vr_path.get()
            mo2_path = self.mo2_path.get()
            
            logging.info(f"update_mo2_configuration called with:")
            logging.info(f"  F4VR path from self.f4vr_path.get(): {f4vr_path}")
            logging.info(f"  MO2 path from self.mo2_path.get(): {mo2_path}")
            
            # Validate paths exist (warn but don't abort if missing)
            f4vr_exe_path = os.path.join(f4vr_path, "Fallout4VR.exe")
            f4sevr_loader_path = os.path.join(f4vr_path, "f4sevr_loader.exe")
            explorer_path = os.path.join(mo2_path, "explorer++", "Explorer++.exe")
           
            if not os.path.exists(f4vr_exe_path):
                logging.warning(f"Fallout4VR.exe not found at {f4vr_exe_path} (may not be critical)")
            if not os.path.exists(f4sevr_loader_path):
                logging.warning(f"f4sevr_loader.exe not found at {f4sevr_loader_path} (may not be installed yet)")
            if not os.path.exists(explorer_path):
                logging.warning(f"Explorer++.exe not found at {explorer_path} (may not be critical)")
            
            # Create file if it doesn't exist
            if not os.path.exists(mo2_ini_path):
                logging.warning(f"MO2 configuration file not found at {mo2_ini_path}, creating new one")
                with open(mo2_ini_path, 'w', encoding='utf-8') as configfile:
                    configfile.write("[General]\ngameName=Fallout 4 VR\n")
                logging.info(f"Created new ModOrganizer.ini at {mo2_ini_path}")
            
            # Read the configuration
            config = configparser.ConfigParser()
            config.read(mo2_ini_path, encoding='utf-8')
            # Convert paths to forward slashes for MO2
            f4vr_path_forward = f4vr_path.replace('\\', '/')
            mo2_path_forward = mo2_path.replace('\\', '/')
            
            logging.info(f"Updating MO2 configuration: F4VR={f4vr_path_forward}, MO2={mo2_path_forward}")
            
            # Update [General] section
            if 'General' not in config:
                config['General'] = {}
            config['General']['gamePath'] = f4vr_path_forward
            config['General']['baseDirectory'] = mo2_path_forward
            
            logging.info(f"Set gamePath to: {f4vr_path_forward}")
            logging.info(f"Set baseDirectory to: {mo2_path_forward}")
            
            # Update [customExecutables] section paths
            if 'customExecutables' not in config:
                config['customExecutables'] = {}
            # Always set Fallout 4 VR executable paths and title
            config['customExecutables']['1\\title'] = "Fallout 4 VR"  # Match the moshortcut name
            config['customExecutables']['1\\binary'] = f"{f4vr_path_forward}/f4sevr_loader.exe"
            config['customExecutables']['1\\workingDirectory'] = f4vr_path_forward
            
            logging.info(f"Set 1\\title to: Fallout 4 VR")
            logging.info(f"Set 1\\binary to: {f4vr_path_forward}/f4sevr_loader.exe")
            logging.info(f"Set 1\\workingDirectory to: {f4vr_path_forward}")
            
            # Always set Explorer++ paths and title
            config['customExecutables']['2\\title'] = "Explorer++"  # Add this line
            config['customExecutables']['2\\binary'] = f"{mo2_path_forward}/explorer++/Explorer++.exe"
            config['customExecutables']['2\\arguments'] = f'"{f4vr_path_forward}/data"'
            config['customExecutables']['2\\workingDirectory'] = f"{mo2_path_forward}/explorer++"
            
            logging.info(f"Set 2\\title to: Explorer++")
            logging.info(f"Set 2\\binary to: {mo2_path_forward}/explorer++/Explorer++.exe")
            logging.info(f"Set 2\\workingDirectory to: {mo2_path_forward}/explorer++")
            
            # Write the updated configuration
            with open(mo2_ini_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            logging.info(f"Successfully updated MO2 configuration file: {mo2_ini_path}")
           
        except Exception as e:
            logging.error(f"Failed to update MO2 configuration: {e}")
            # Don't raise - this is not critical for installation

    def copy_frik_ini(self):
        """Copy FRIK weapon offsets for Fallout London VR"""
        try:
            self.root.after(0, lambda: self.message_label.config(text="Copying FRIK weapon offsets", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.debug("Starting FRIK weapon offsets copy")
            
            # Copy weapon offsets files
            self.copy_weapon_offsets()
            
            self.root.after(0, lambda: self.message_label.config(text="FRIK weapon offsets copied successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("FRIK weapon offsets copy completed successfully")
            
        except Exception as e:
            logging.error(f"Failed to copy FRIK weapon offsets: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to copy FRIK weapon offsets: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)

    def copy_weapon_offsets(self):
        """Copy all weapon offset files from assets to user's FRIK_Config"""
        try:
            # Source directory - try multiple locations
            # 1. Try bundled assets first (for fresh installs from standalone assets)
            assets_dir = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "MO2", "mods", "Fallout London VR", "F4SE", "Plugins", "FRIK_weapon_offsets")
            
            # 2. If not found and we have an MO2 path, try the installed location (for updates)
            if not os.path.exists(assets_dir) and hasattr(self, 'mo2_path') and self.mo2_path.get():
                installed_dir = os.path.join(self.mo2_path.get(), "mods", "Fallout London VR", "F4SE", "Plugins", "FRIK_weapon_offsets")
                if os.path.exists(installed_dir):
                    assets_dir = installed_dir
                    logging.info(f"Using weapon offsets from installed mod directory: {assets_dir}")
            
            # Destination directory in user's Documents
            documents_folder = os.path.expanduser("~\\Documents")
            frik_config_dir = os.path.join(documents_folder, "My Games", "Fallout4VR", "FRIK_Config")
            dest_dir = os.path.join(frik_config_dir, "Weapons_Offsets")
            os.makedirs(dest_dir, exist_ok=True)
            
            # Copy FRIK_FOLVR.ini from assets root to FRIK_Config (without overwrite)
            assets_root = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets")
            frik_folvr_src = os.path.join(assets_root, "FRIK_FOLVR.ini")
            frik_folvr_dest = os.path.join(frik_config_dir, "FRIK_FOLVR.ini")
            
            if os.path.exists(frik_folvr_src):
                if not os.path.exists(frik_folvr_dest):
                    shutil.copy2(frik_folvr_src, frik_folvr_dest)
                    logging.info(f"Copied FRIK_FOLVR.ini to {frik_folvr_dest}")
                else:
                    logging.info(f"FRIK_FOLVR.ini already exists at {frik_folvr_dest}, skipping (no overwrite)")
            else:
                logging.warning(f"FRIK_FOLVR.ini not found at {frik_folvr_src}")
            
            if not os.path.exists(assets_dir):
                logging.warning(f"Weapon offsets source directory does not exist: {assets_dir}")
                logging.warning("Skipping weapon offsets copy - directory not found in bundled assets or installed location")
                return
            
            files = os.listdir(assets_dir)
            if not files:
                logging.warning(f"Weapon offsets directory is empty: {assets_dir}")
                return
                
            for file in files:
                src_file = os.path.join(assets_dir, file)
                dest_file = os.path.join(dest_dir, file)
                if os.path.isfile(src_file):
                    shutil.copy2(src_file, dest_file)
                    logging.info(f"Copied weapon offset file: {file}")
            
            logging.info(f"Copied {len(files)} weapon offset files from {assets_dir}")
        except Exception as e:
            logging.error(f"Failed to copy weapon offset files: {e}")

    def install_xse_plugin_preloader(self):
        """Install xSE Plugin Preloader to F4VR installation directory.
        
        Copies WinHTTP.dll and xSE PluginPreloader.xml to F4VR root.
        Creates backups of existing files before overwriting.
        """
        try:
            self.root.after(0, lambda: self.message_label.config(text="Installing xSE Plugin Preloader...", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("Starting xSE Plugin Preloader installation")
            
            # Get F4VR installation path
            f4vr_path = self.f4vr_path.get()
            if not f4vr_path or not os.path.exists(f4vr_path):
                logging.warning("F4VR path not set or does not exist, skipping xSE Plugin Preloader installation")
                return False
            
            # Get source files from assets
            assets_dir = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets")
            
            # Source files
            src_winhttp = os.path.join(assets_dir, "WinHTTP.dll")
            src_xml = os.path.join(assets_dir, "xSE PluginPreloader.xml")
            
            # Check source files exist
            if not os.path.exists(src_winhttp):
                logging.error(f"Source WinHTTP.dll not found at {src_winhttp}")
                return False
            if not os.path.exists(src_xml):
                logging.error(f"Source xSE PluginPreloader.xml not found at {src_xml}")
                return False
            
            # Destination files
            dest_winhttp = os.path.join(f4vr_path, "WinHTTP.dll")
            dest_xml = os.path.join(f4vr_path, "xSE PluginPreloader.xml")
            
            # Backup existing WinHTTP.dll if it exists
            if os.path.exists(dest_winhttp):
                backup_path = os.path.join(f4vr_path, "WinHTTP.dll.folvrbackup")
                if not os.path.exists(backup_path):
                    shutil.copy2(dest_winhttp, backup_path)
                    logging.info(f"Created backup of WinHTTP.dll at {backup_path}")
                else:
                    logging.info(f"Backup already exists at {backup_path}")
            
            # Backup existing xSE PluginPreloader.xml if it exists
            if os.path.exists(dest_xml):
                backup_xml_path = os.path.join(f4vr_path, "xSE PluginPreloader.xml.folvrbackup")
                if not os.path.exists(backup_xml_path):
                    shutil.copy2(dest_xml, backup_xml_path)
                    logging.info(f"Created backup of xSE PluginPreloader.xml at {backup_xml_path}")
                else:
                    logging.info(f"Backup already exists at {backup_xml_path}")
            
            # Copy files to F4VR root
            shutil.copy2(src_winhttp, dest_winhttp)
            logging.info(f"Copied WinHTTP.dll to {dest_winhttp}")
            
            shutil.copy2(src_xml, dest_xml)
            logging.info(f"Copied xSE PluginPreloader.xml to {dest_xml}")
            
            self.root.after(0, lambda: self.message_label.config(text="xSE Plugin Preloader installed", fg="#ffffff") if self.message_label.winfo_exists() else None)
            logging.info("xSE Plugin Preloader installation completed successfully")
            return True
            
        except Exception as e:
            logging.error(f"Failed to install xSE Plugin Preloader: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"xSE Plugin Preloader failed: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            return False

    def initialize_and_start_slideshow(self):
        """Initialize and start the slideshow after install button is clicked"""
        try:
            logging.info("=== INITIALIZING SLIDESHOW ===")
            
            # Create slideshow if not already created
            if not hasattr(self, 'slideshow') or self.slideshow is None:
                # Find the content frame or main container
                content_parent = None
                
                # First try to find the main container (should exist)
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Frame):
                        # Check if this is the main container frame
                        for child in widget.winfo_children():
                            if isinstance(child, tk.Frame):
                                # This is likely our content frame
                                content_parent = child
                                logging.info(f"Found content frame: {content_parent}")
                                break
                        if content_parent:
                            break
                
                # If we didn't find a nested frame, use the first frame we find
                if not content_parent:
                    for widget in self.root.winfo_children():
                        if isinstance(widget, tk.Frame):
                            content_parent = widget
                            logging.info(f"Using first frame found: {content_parent}")
                            break
                
                # Final fallback to root
                if not content_parent:
                    content_parent = self.root
                    logging.warning("No frame found, using root window for slideshow")
                
                logging.info(f"Parent selected for slideshow: {content_parent}")
                logging.info(f"Parent type: {type(content_parent)}")
                logging.info(f"Parent exists: {content_parent.winfo_exists()}")
                
                # Create the slideshow instance
                self.slideshow = SlideshowFrame(content_parent, self)
                logging.info("Slideshow instance created")
            else:
                logging.info("Slideshow already exists")
            
            # Start the slideshow
            if self.slideshow:
                self.slideshow.start_slideshow()
                logging.info("Slideshow start method called")
            else:
                logging.error("Slideshow instance is None!")
            
        except Exception as e:
            logging.error(f"Failed to start slideshow: {e}")
            import traceback
            logging.error(traceback.format_exc())
            # Non-critical feature, continue without slideshow

    def show_completion_ui(self):
        """Show completion screen"""
        # Stop slideshow if running
        if hasattr(self, 'slideshow') and self.slideshow:
            self.slideshow.stop_slideshow()
        
        # Hide window to prevent pop-in
        self.root.withdraw()
        # Destroy existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        # Recreate the custom title bar
        self.add_custom_title_bar()
        bg_color = "#1e1e1e"
        fg_color = "#ffffff"
        accent_color = "#0078d7"
        if self.bg_image:
            bg_label = tk.Label(self.root, image=self.bg_image)
            bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            bg_label.lower()
            main_container = tk.Frame(self.root)
        else:
            self.root.configure(bg=bg_color)
            main_container = tk.Frame(self.root, bg=bg_color)
        main_container.pack(fill="both", expand=True, padx=self.get_scaled_value(20), pady=self.get_scaled_value(5))
        
        frame_bg = bg_color if not self.bg_image else '#1e1e1e'
        
        # Create a canvas with scrollbar for small screens
        canvas = tk.Canvas(main_container, bg=frame_bg, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        # Create scrollable frame inside canvas
        scrollable_frame = tk.Frame(canvas, bg=frame_bg)
        
        # Configure canvas scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="n")
        
        # Make the scrollable frame expand to canvas width
        def configure_scroll_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', configure_scroll_width)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas (scrollbar hidden but functional)
        canvas.pack(side="left", fill="both", expand=True)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Enable arrow key scrolling
        def on_key_up(event):
            canvas.yview_scroll(-1, "units")
        def on_key_down(event):
            canvas.yview_scroll(1, "units")
        self.root.bind("<Up>", on_key_up)
        self.root.bind("<Down>", on_key_down)
        
        # Store canvas reference
        self.completion_canvas = canvas
        
        # Content goes in scrollable_frame
        content_frame = scrollable_frame
        
        # Show logo at the top
        if self.logo:
            tk.Label(content_frame, image=self.logo, bg=frame_bg).pack(pady=(self.get_scaled_value(5), self.get_scaled_value(5)))
        
        # Different text based on installation mode
        if self.update_mode:
            title_text = "Update Complete!"
            profiles_path = os.path.join(self.mo2_path.get(), "profiles", "Default")
            message_text = f"Fallout: London VR has been successfully updated to 0.99.\n\nYour .ini files have been backed up to:\n{profiles_path}"
        else:
            title_text = "Installation Complete!"
            message_text = "Fallout: London VR has been successfully installed.\n\nA desktop shortcut and a Fallout London VR start menu folder have been created. If you want to add more mods, you can start Modorganizer through the start menu shortcut."
        
        tk.Label(content_frame, text=title_text, font=self.title_font, bg=frame_bg, fg="#00ff00").pack(pady=(self.get_scaled_value(5), self.get_scaled_value(5)))
        
        tk.Label(content_frame, text=message_text,
                 font=self.regular_font, bg=frame_bg, fg=fg_color, wraplength=self.get_scaled_value(460), justify="center").pack(pady=(self.get_scaled_value(5), self.get_scaled_value(10)))
        
        # Show Atkins image
        if hasattr(self, 'atkins_image') and self.atkins_image:
            # Create a larger version of atkins for completion screen (15% bigger than welcome)
            try:
                atkins_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "atkins.png")
                if os.path.exists(atkins_path):
                    atkins_img = Image.open(atkins_path)
                    # 15% bigger than welcome screen (310 * 1.15 = 357)
                    atkins_width = self.get_scaled_value(357)
                    aspect_ratio = atkins_img.height / atkins_img.width
                    atkins_height = int(atkins_width * aspect_ratio)
                    atkins_img = atkins_img.resize((atkins_width, atkins_height), Image.Resampling.LANCZOS)
                    self.completion_atkins = ImageTk.PhotoImage(atkins_img)
                    tk.Label(content_frame, image=self.completion_atkins, bg=frame_bg).pack(pady=(self.get_scaled_value(5), self.get_scaled_value(10)))
            except Exception as e:
                logging.warning(f"Could not load atkins for completion screen: {e}")
        
        # Button frame inside scrollable content
        button_frame = tk.Frame(content_frame, bg=frame_bg)
        button_frame.pack(pady=(self.get_scaled_value(10), self.get_scaled_value(20)))
        
        tk.Button(button_frame, text="Finish Installation", command=self.exit_installer,
                 font=self.bold_font, bg=accent_color, fg=fg_color, bd=0, relief="flat",
                 activebackground="#005ba1", padx=self.get_scaled_value(20), pady=self.get_scaled_value(10), width=18).pack(side=tk.LEFT, padx=self.get_scaled_value(5))
        
        tk.Button(button_frame, text="Launch Game", command=self.launch_game_from_completion,
                 font=self.bold_font, bg="#00aa00", fg=fg_color, bd=0, relief="flat",
                 activebackground="#008800", padx=self.get_scaled_value(20), pady=self.get_scaled_value(10), width=18).pack(side=tk.LEFT, padx=self.get_scaled_value(5))
        
        # Force update and show window
        self.root.update_idletasks()
        self.root.deiconify()
        
        # Open donation pages in background when completion UI is shown
    # self.open_donation_pages_background()  # Disabled: do not open donation page at end of installation

    def launch_game_from_completion(self):
        """Launch the game directly from the completion screen.
        
        We create a temporary batch file that waits briefly then launches MO2.
        This allows the installer to fully exit before MO2 starts, preventing
        the 'failed to remove temporary directory' error.
        """
        try:
            mo2_exe = os.path.join(self.mo2_path.get(), "ModOrganizer.exe")
            
            if not os.path.exists(mo2_exe):
                messagebox.showerror("Error", f"ModOrganizer.exe not found at {mo2_exe}")
                logging.error(f"MO2 executable not found at {mo2_exe}")
                return
            
            logging.info(f"Launching game from completion screen via batch launcher")
            
            # Create a temporary VBScript that will:
            # 1. Wait 2 seconds for installer to fully exit
            # 2. Launch MO2 silently (no window)
            # 3. Delete itself
            vbs_content = f'''Set WshShell = CreateObject("WScript.Shell")
WScript.Sleep 2000
WshShell.CurrentDirectory = "{self.mo2_path.get()}"
WshShell.Run """{mo2_exe}"" ""moshortcut://Portable:Fallout 4 VR""", 0, False
Set fso = CreateObject("Scripting.FileSystemObject")
fso.DeleteFile WScript.ScriptFullName
'''
            
            # Write VBS file to user's temp directory (not PyInstaller's temp)
            user_temp = os.environ.get('TEMP', os.environ.get('TMP', 'C:\\Windows\\Temp'))
            vbs_path = os.path.join(user_temp, 'launch_fallout_london_vr.vbs')
            
            with open(vbs_path, 'w') as f:
                f.write(vbs_content)
            
            # Launch the VBScript silently using wscript
            CREATE_NO_WINDOW = 0x08000000
            
            subprocess.Popen(
                ['wscript.exe', vbs_path],
                creationflags=CREATE_NO_WINDOW,
                close_fds=True
            )
            
            logging.info("VBScript launcher started, closing installer")
            
            # Close the installer immediately
            self.root.destroy()
            os._exit(0)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch game: {e}")
            logging.error(f"Failed to launch game: {e}")

    def open_donation_pages_background(self):
        """Open donation page in background without stealing focus"""
        urls = [
            ("Fallout London VR Donation", "https://ko-fi.com/falloutlondonvr")
        ]
        
        try:
            for name, url in urls:
                try:
                    if sys.platform == 'win32':
                        # Windows: Use cmd /c start "" /min <url> so the new window is created minimized.
                        # Use creationflags to avoid opening an extra console window when running from GUI.
                        try:
                            subprocess.Popen([
                                'cmd', '/c', 'start', '""', '/min', url
                            ], shell=False, creationflags=subprocess.CREATE_NO_WINDOW)
                        except TypeError:
                            # Older Python on some environments might not accept creationflags; fall back
                            subprocess.Popen(['cmd', '/c', 'start', '""', '/min', url], shell=False)
                    else:
                        # Other platforms: try to open without raising the browser window
                        webbrowser.open(url, new=2, autoraise=False)
                    
                    logging.info(f"Opened {name} URL in background: {url}")
                    time.sleep(0.2)  # Small delay between opens
                    
                except Exception as e:
                    logging.warning(f"Failed to open {name} URL: {e}")
                    # Continue with other URLs even if one fails
            
            logging.info("Donation pages opened in background")
            
        except Exception as e:
            logging.warning(f"Error opening donation pages: {e}")
            # Non-critical error, don't show to user

    def exit_installer(self):
        """Close the installer application"""
        try:
            # Cleanup slideshow
            if hasattr(self, 'slideshow') and self.slideshow:
                self.slideshow.cleanup()
            
            logging.info("Exiting installer")
            sys.exit(0)
        except Exception as e:
            logging.error(f"Failed to exit installer: {e}")
            sys.exit(0)

    def launch_mo2(self):
        """Launch MO2, ensure Steam VR is running, and retry if game doesn't start"""
        try:
            # First, open donation pages in background
            # self.open_donation_pages_background()  # Disabled: do not open donation page at end of installation
            
            mo2_exe = os.path.join(self.mo2_path.get(), "ModOrganizer.exe")
            
            if not os.path.exists(mo2_exe):
                messagebox.showerror("Error", f"ModOrganizer.exe not found at {mo2_exe}")
                logging.error(f"MO2 executable not found at {mo2_exe}")
                return
            
            # Check if Steam VR is running
            steam_vr_running = any("vrserver.exe" in proc.name().lower() for proc in psutil.process_iter(['name']))
            logging.info(f"Steam VR running before launch: {steam_vr_running}")
            
            if not steam_vr_running:
                # Attempt to start Steam VR
                logging.info("Starting Steam VR via steam://run/250820")
                subprocess.Popen(["start", "steam://run/250820"], shell=True)
                
                # Wait for Steam VR to initialize (up to 10 seconds)
                start_time = time.time()
                while time.time() - start_time < 10:
                    steam_vr_running = any("vrserver.exe" in proc.name().lower() for proc in psutil.process_iter(['name']))
                    if steam_vr_running:
                        logging.info("Steam VR started successfully")
                        break
                    time.sleep(1)
                else:
                    logging.warning("Steam VR did not start within 10 seconds")
            
            # Minimize the Tkinter window
            self.minimize_window(self.root)
            logging.info("Minimized Tkinter window before launching MO2")
            
            # Launch MO2 (first attempt)
            cmd = [mo2_exe, "moshortcut://Portable:Fallout 4 VR"]
            logging.info(f"Launching MO2 (first attempt) with command: {' '.join(cmd)}")
            
            process = subprocess.Popen(
                cmd,
                cwd=self.mo2_path.get()
            )
            logging.info(f"MO2 process started with PID: {process.pid}")
            
            # Wait 10 seconds and check if Fallout4VR.exe is running
            time.sleep(10)
            game_running = any("fallout4vr.exe" in proc.name().lower() for proc in psutil.process_iter(['name']))
            logging.info(f"Fallout 4 VR running after first attempt: {game_running}")
            
            if not game_running:
                # Retry launching MO2
                logging.info(f"Launching MO2 (second attempt) with command: {' '.join(cmd)}")
                process = subprocess.Popen(
                    cmd,
                    cwd=self.mo2_path.get()
                )
                logging.info(f"MO2 process started with PID: {process.pid} (second attempt)")
            
            # Wait longer before closing installer to allow Steam, Steam VR, and MO2 initialization
            self.root.after(15000, lambda: sys.exit(0))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch MO2: {e}")
            logging.error(f"Failed to launch MO2: {e}")

    def download_with_memory_management(self, url, output_path, label_text, verify_ssl=True):
        """Generic download function with proper memory management and retry logic"""
        max_retries = 5
        retry_delay = 5  # seconds between retries
        
        self.create_progress_bar(f"Downloading {label_text}")
        
        for attempt in range(1, max_retries + 1):
            try:
                if attempt > 1:
                    self.root.after(0, lambda a=attempt, lt=label_text: self.progress_label.config(
                        text=f"Retrying {lt} download (attempt {a}/{max_retries})..."))
                    self.root.after(0, lambda: self.message_label.config(
                        text="Trying to reconnect. Please check your internet connection.", fg="#ffaa00") if self.message_label.winfo_exists() else None)
                    time.sleep(retry_delay)
                
                with requests.get(url, stream=True, verify=verify_ssl, timeout=30) as r:
                    r.raise_for_status()
                    
                    # Connection restored - clear the warning message
                    if attempt > 1:
                        self.root.after(0, lambda: self.message_label.config(
                            text="Connection restored. Resuming download", fg="#00ff00") if self.message_label.winfo_exists() else None)
                    
                    total_size = int(r.headers.get('content-length', 0))
                    downloaded_size = 0
                    chunk_size = 256 * 1024  # 256KB chunks for downloads
                    
                    with open(output_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=chunk_size):
                            if self.cancel_requested:
                                raise InterruptedError("Installation cancelled by user")
                            
                            f.write(chunk)
                            downloaded_size += len(chunk)
                            
                            # Periodic flush to disk
                            if downloaded_size % (10 * 1024 * 1024) == 0:  # Every 10MB
                                f.flush()
                                os.fsync(f.fileno())
                            
                            # Update progress
                            progress_percentage = (downloaded_size / total_size) * 100 if total_size > 0 else 0
                            if self.progress and self.progress.winfo_exists():
                                self.root.after(0, lambda pp=progress_percentage: 
                                    self.progress.__setitem__("value", pp) if self.progress.winfo_exists() else None)
                            if self.progress_label and self.progress_label.winfo_exists():
                                self.root.after(0, lambda pp=progress_percentage, lt=label_text: 
                                    self.progress_label.config(text=f"Downloading {lt} ({pp:.1f}%)") 
                                    if self.progress_label.winfo_exists() else None)
                
                logging.info(f"Downloaded {label_text} to {output_path}")
                return output_path
                
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout, 
                    requests.exceptions.ChunkedEncodingError) as e:
                logging.warning(f"Download attempt {attempt}/{max_retries} for {label_text} failed: {e}")
                
                if attempt < max_retries:
                    self.root.after(0, lambda a=attempt, d=retry_delay, lt=label_text: self.progress_label.config(
                        text=f"Connection failed. Retrying {lt} in {d}s (attempt {a}/{max_retries})..."))
                    self.root.after(0, lambda: self.message_label.config(
                        text="Connection error. Please check your internet connection.", fg="#ffaa00") if self.message_label.winfo_exists() else None)
                else:
                    error_msg = f"Failed to download {label_text}: Connection error after multiple attempts.\nPlease check your internet connection and try again."
                    self.root.after(0, lambda m=error_msg: self.message_label.config(text=m, fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Failed to download {label_text} after {max_retries} attempts: {e}")
                    raise Exception(error_msg)
                    
            except requests.exceptions.HTTPError as e:
                error_msg = f"Failed to download {label_text}: Server error {e.response.status_code}"
                self.root.after(0, lambda m=error_msg: self.message_label.config(text=m, fg="#ff6666") if self.message_label.winfo_exists() else None)
                logging.error(f"Failed to download {label_text}: {e}")
                raise Exception(error_msg)
                
            except InterruptedError:
                raise
                
            except Exception as e:
                if attempt < max_retries:
                    logging.warning(f"Download attempt {attempt}/{max_retries} for {label_text} failed: {e}")
                    continue
                else:
                    error_msg = f"Failed to download {label_text}: {type(e).__name__}"
                    self.root.after(0, lambda m=error_msg: self.message_label.config(text=m, fg="#ff6666") if self.message_label.winfo_exists() else None)
                    logging.error(f"Failed to download {label_text}: {e}")
                    raise Exception(error_msg)

    def download_mo2_portable(self):
        """Download portable MO2 archive for inline installation"""
        mo2_url = "https://github.com/ModOrganizer2/modorganizer/releases/download/v2.5.2/Mod.Organizer-2.5.2.7z"
        temp_dir = os.path.join(tempfile.gettempdir())
        mo2_archive = os.path.join(temp_dir, "Mod.Organizer-2.5.2.7z")

        self.root.after(0, lambda: self.message_label.config(text="Setting up Mod Organizer 2", fg="#ffffff") if self.message_label.winfo_exists() else None)
        
        # Use the new generic download function
        return self.download_with_memory_management(mo2_url, mo2_archive, "MO2")

    def extract_mo2(self, archive_path):
        """Extract MO2 and configure for portable mode"""
        mo2_extract_dir = self.mo2_path.get()  # Extract directly to install dir
        os.makedirs(mo2_extract_dir, exist_ok=True)

        self.root.after(0, lambda: self.message_label.config(text="Setting up Mod Organizer 2", fg="#ffffff") if self.message_label.winfo_exists() else None)
        self.create_progress_bar("Extracting MO2")

        try:
            # Use bundled 7za.exe instead of py7zr
            bundled_7za = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "7za.exe")
            
            if not os.path.exists(bundled_7za):
                raise FileNotFoundError(f"Bundled 7za.exe not found at {bundled_7za}")

            # Extract using bundled 7za
            self.root.after(0, lambda: self.progress_label.config(text="Extracting MO2"))
            extract_cmd = [bundled_7za, "x", archive_path, f"-o{mo2_extract_dir}", "-y", "-bb0", "-bd"]
            
            logging.info(f"Extracting MO2 using bundled 7za: {' '.join(extract_cmd)}")
            result = subprocess.run(extract_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if result.returncode != 0:
                logging.error(f"7za stdout: {result.stdout}")
                logging.error(f"7za stderr: {result.stderr}")
                raise Exception(f"7za extraction failed with code {result.returncode}: {result.stderr}")
            
            logging.info(f"7za extraction successful")
            
            # Update progress to 100%
            self.root.after(0, lambda: self.progress.__setitem__("value", 100))
            self.root.after(0, lambda: self.progress_label.config(text="Extracting MO2 (100%)"))

            # Configure for portable mode
            mo2_exe_path = os.path.join(mo2_extract_dir, "ModOrganizer.exe")
            if os.path.exists(mo2_exe_path):
                # Create portable mode indicator
                portable_file = os.path.join(mo2_extract_dir, "portable.txt")
                with open(portable_file, 'w') as f:
                    f.write("This file enables portable mode for MO2")
                
                logging.info("MO2 configured for portable mode")
            else:
                logging.warning(f"ModOrganizer.exe not found at {mo2_exe_path} after extraction")

            logging.info(f"Extracted MO2 to {mo2_extract_dir}")
            self.root.after(0, lambda: self.message_label.config(text="Mod Organizer 2 installed successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Clean up downloaded archive
            if os.path.exists(archive_path):
                os.unlink(archive_path)
                
        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to extract MO2: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            logging.error(f"Failed to extract MO2: {e}")
            raise

    def create_progress_bar(self, label_text):
        """Create a progress bar with label"""
        # Destroy existing progress bar and label if they exist
        if hasattr(self, 'progress') and self.progress and self.progress.winfo_exists():
            self.progress.destroy()
        if hasattr(self, 'progress_label') and self.progress_label and self.progress_label.winfo_exists():
            self.progress_label.destroy()

        # Create a frame to hold the progress bar and label and place it reliably above the bottom
        if hasattr(self, 'progress_frame') and self.progress_frame and self.progress_frame.winfo_exists():
            self.progress_frame.destroy()
        # Ensure geometry is calculated
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        self.progress_frame = tk.Frame(self.root, bg="#1e1e1e", width=getattr(self, 'window_width', None))

        # Try to center the progress frame in the free space between the slideshow and the window bottom
        desired_y = None
        try:
            # Ensure geometry metrics are up to date
            self.root.update_idletasks()
            root_h = self.root.winfo_height()
            root_rooty = self.root.winfo_rooty()

            # Prefer the image_label for precise image bottom if available
            if hasattr(self, 'slideshow') and self.slideshow and hasattr(self.slideshow, 'image_label') and self.slideshow.image_label.winfo_exists():
                il = self.slideshow.image_label
                il.update_idletasks()
                il_h = il.winfo_height()
                il_rooty = il.winfo_rooty()
                il_y_rel = max(0, il_rooty - root_rooty)
                bottom_of_images = il_y_rel + il_h

                # Compute available vertical space below images
                free_space_top = bottom_of_images
                free_space_bottom = root_h
                available = free_space_bottom - free_space_top
                # Only center if there's a meaningful free space
                if available > 40:
                    desired_y = int(free_space_top + available / 2)
            else:
                # Fallback to slideshow_frame if image_label not available
                if hasattr(self, 'slideshow') and self.slideshow and hasattr(self.slideshow, 'slideshow_frame') and self.slideshow.slideshow_frame.winfo_exists():
                    sf = self.slideshow.slideshow_frame
                    sf.update_idletasks()
                    sf_h = sf.winfo_height()
                    sf_rooty = sf.winfo_rooty()
                    sf_y_rel = max(0, sf_rooty - root_rooty)
                    bottom_of_slideshow = sf_y_rel + sf_h
                    available = root_h - bottom_of_slideshow
                    if available > 40:
                        desired_y = int(bottom_of_slideshow + available / 2)
        except Exception as e:
            logging.warning(f"Progress placement computation failed: {e}")
            desired_y = None

        # Place the frame. If we computed a desired y, anchor at center at that pixel y; otherwise default to 38px above bottom
        try:
            if desired_y is not None:
                self.progress_frame.place(relx=0.5, y=desired_y, anchor='center')
            else:
                self.progress_frame.place(relx=0.5, rely=1.0, anchor='s', y=-38)
        except Exception:
            # Fallback to previous placement on any error
            try:
                self.progress_frame.place(relx=0.5, rely=1.0, anchor='s', y=-38)
            except Exception:
                pass

        # Ensure the progress frame is above other packed widgets so it's visible
        try:
            self.progress_frame.lift()
            self.progress_frame.tkraise()
        except Exception:
            pass

        # Create progress label centered in the frame
        self.progress_label = tk.Label(
            self.progress_frame,
            text=label_text,
            font=self.bold_font,
            bg="#1e1e1e",
            fg="#ffffff"
        )
        self.progress_label.pack(pady=6)

        # Create progress bar centered in the frame
        self.progress = ttk.Progressbar(self.progress_frame, length=300, mode="determinate")
        self.progress.pack(pady=6)
        self.root.update()
        return self.progress

    def download_f4sevr(self):
        """Download F4SEVR archive from official source with retry logic"""
        f4sevr_url = "https://f4se.silverlock.org/beta/f4sevr_0_6_21.7z"
        temp_dir = os.path.join(tempfile.gettempdir())
        f4sevr_archive = os.path.join(temp_dir, "f4sevr_0_6_21.7z")

        self.root.after(0, lambda: self.message_label.config(text="Setting up Fallout 4 Script Extender VR", fg="#ffffff") if self.message_label.winfo_exists() else None)
        
        # Download F4SEVR using the retry-enabled download function
        # Note: verify_ssl=False because f4se.silverlock.org has a weak certificate
        self.download_with_memory_management(f4sevr_url, f4sevr_archive, "F4SEVR", verify_ssl=False)
        logging.info(f"Downloaded F4SEVR archive to {f4sevr_archive}")
        return f4sevr_archive

    def extract_and_install_f4sevr(self, archive_path, dest_dir):
        """Extract F4SEVR archive and install to Fallout 4 VR directory"""
        temp_extract_dir = os.path.join(tempfile.gettempdir(), "f4sevr_extract")
        
        self.root.after(0, lambda: self.message_label.config(text="Setting up Fallout 4 Script Extender VR", fg="#ffffff") if self.message_label.winfo_exists() else None)
        self.create_progress_bar("Extracting F4SEVR")

        try:
            # Clean up any existing temp directory
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir, ignore_errors=True)
            
            os.makedirs(temp_extract_dir, exist_ok=True)

            # Use bundled 7za.exe
            bundled_7za = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "7za.exe")
            
            if not os.path.exists(bundled_7za):
                raise FileNotFoundError(f"Bundled 7za.exe not found at {bundled_7za}")

            # Extract using bundled 7za
            self.root.after(0, lambda: self.progress_label.config(text="Extracting F4SEVR"))
            extract_cmd = [bundled_7za, "x", archive_path, f"-o{temp_extract_dir}", "-y", "-bb0", "-bd"]
            
            result = subprocess.run(extract_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode != 0:
                raise Exception(f"7za extraction failed: {result.stderr}")
            
            logging.info(f"Extracted F4SEVR using bundled 7za to {temp_extract_dir}")

            # Update progress to 50%
            self.root.after(0, lambda: self.progress.__setitem__("value", 50))
            self.root.after(0, lambda: self.progress_label.config(text="Installing F4SEVR files"))

            # Now copy files from temp directory to F4VR directory
            self.copy_f4sevr_files(temp_extract_dir, dest_dir)

            # Clean up temporary directory
            shutil.rmtree(temp_extract_dir, ignore_errors=True)
            
            # Clean up downloaded archive
            if os.path.exists(archive_path):
                os.unlink(archive_path)

            logging.info("F4SEVR installation completed successfully")
            self.root.after(0, lambda: self.message_label.config(text="F4SEVR installed successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)

        except Exception as e:
            # Clean up on error
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir, ignore_errors=True)
            
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to install F4SEVR: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            logging.error(f"F4SEVR installation failed: {e}")
            raise

    def copy_f4sevr_files(self, src_dir, dest_dir):
        """Copy F4SEVR files from extracted directory to Fallout 4 VR root directory"""
        try:
            if not os.path.exists(src_dir):
                raise FileNotFoundError(f"F4SEVR source directory not found: {src_dir}")

            # Check if there's a f4sevr_0_6_21 folder inside the extracted directory
            f4sevr_folder = None
            for item in os.listdir(src_dir):
                item_path = os.path.join(src_dir, item)
                if os.path.isdir(item_path) and item.startswith("f4sevr"):
                    f4sevr_folder = item_path
                    break
            
            # If we found the f4sevr folder, use its contents as the source
            if f4sevr_folder:
                actual_src_dir = f4sevr_folder
                logging.info(f"Found F4SEVR folder: {f4sevr_folder}, using its contents")
            else:
                actual_src_dir = src_dir
                logging.info("No F4SEVR subfolder found, using extraction directory directly")

            total_size = 0
            file_count = 0
            files_to_copy = []
            
            # Calculate total size and build file list from the actual source directory
            for dirpath, dirnames, filenames in os.walk(actual_src_dir):
                for filename in filenames:
                    src_file = os.path.join(dirpath, filename)
                    
                    # Calculate relative path from the actual source directory
                    rel_path = os.path.relpath(dirpath, actual_src_dir)
                    
                    # Create corresponding path in destination
                    if rel_path == '.':
                        dest_file = os.path.join(dest_dir, filename)
                    else:
                        dest_file = os.path.join(dest_dir, rel_path, filename)
                    
                    try:
                        file_size = os.path.getsize(src_file)
                        total_size += file_size
                        file_count += 1
                        files_to_copy.append((src_file, dest_file, file_size))
                    except OSError as e:
                        logging.warning(f"Could not get size of {src_file}: {e}")

            logging.info(f"F4SEVR files prepared for copying")
            
            if file_count == 0:
                logging.warning(f"No files found to copy from {actual_src_dir}")
                return

            # Rest of the method remains the same...
            copied_size = 0
            files_copied = 0

            for src_file, dest_file, file_size in files_to_copy:
                if self.cancel_requested:
                    raise InterruptedError("Installation cancelled by user")

                # Create destination directory if it doesn't exist
                dest_file_dir = os.path.dirname(dest_file)
                os.makedirs(dest_file_dir, exist_ok=True)

                try:
                    # Handle read-only files in destination
                    if os.path.exists(dest_file):
                        self.remove_readonly_and_overwrite(dest_file)
                    
                    shutil.copy2(src_file, dest_file)
                    copied_size += file_size
                    files_copied += 1

                except PermissionError as e:
                    logging.warning(f"Permission error copying {src_file} to {dest_file}: {e}")
                    # Try to handle read-only file
                    try:
                        self.remove_readonly_and_overwrite(dest_file)
                        shutil.copy2(src_file, dest_file)
                        copied_size += file_size
                        files_copied += 1
                        logging.info(f"Successfully copied after removing read-only: {dest_file}")
                    except Exception as retry_error:
                        logging.error(f"Failed to copy {src_file} after retry: {retry_error}")
                        # Continue with other files instead of failing completely
                        continue

                # Update progress (50% offset from extraction)
                progress_percentage = 50 + (copied_size / total_size) * 50
                self.root.after(0, lambda pp=progress_percentage: self.progress.__setitem__("value", pp))
                self.root.after(0, lambda pp=progress_percentage: self.progress_label.config(text=f"Installing F4SEVR ({pp:.1f}%)"))

            logging.info(f"F4SEVR file copy completed: {files_copied} files copied")

        except Exception as e:
            logging.error(f"F4SEVR file copy failed: {e}")
            raise

    def download_and_install_frik(self):
        """Download and install FRIK to mods directory"""
        try:
            frik_url = "https://github.com/rollingrock/Fallout-4-VR-Body/releases/download/v0.76/FRIK.-.v0.76.10.-.20251201.7z"
            temp_dir = os.path.join(tempfile.gettempdir())
            frik_archive = os.path.join(temp_dir, "FRIK.v0.76.10.7z")
            
            self.root.after(0, lambda: self.message_label.config(text="Updating FRIK VR Body", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Download FRIK using the retry-enabled download function
            self.download_with_memory_management(frik_url, frik_archive, "FRIK")
            logging.info(f"Downloaded FRIK archive to {frik_archive}")
            
            # Extract FRIK
            self.extract_frik(frik_archive)
            
        except Exception as e:
            logging.error(f"FRIK installation failed: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to install FRIK: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def extract_frik(self, archive_path):
        """Extract FRIK to mods directory"""
        try:
            # Create FRIK mod directory
            frik_mod_dir = os.path.join(self.mo2_path.get(), "mods", "FRIK")
            os.makedirs(frik_mod_dir, exist_ok=True)
            
            self.root.after(0, lambda: self.message_label.config(text="Extracting FRIK VR Body", fg="#ffffff") if self.message_label.winfo_exists() else None)
            self.create_progress_bar("Extracting FRIK")
            
            # Extract using py7zr since it's a .7z file
            logging.info(f"Extracting FRIK 7z archive: {archive_path}")
            logging.info(f"Archive size: {os.path.getsize(archive_path)} bytes")
            
            try:
                # Extract all files at once instead of one by one
                with py7zr.SevenZipFile(archive_path, mode='r') as z:
                    # Extract all files
                    z.extractall(path=frik_mod_dir)
                    
                # Update progress to 100%
                self.root.after(0, lambda: self.progress.__setitem__("value", 100))
                self.root.after(0, lambda: self.progress_label.config(text=f"Extracting FRIK (100%)"))
            
            except Exception as extract_error:
                # Log detailed error information
                logging.error(f"Archive extraction failed: {extract_error}")
                logging.error(f"Archive path: {archive_path}")
                logging.error(f"Archive exists: {os.path.exists(archive_path)}")
                if os.path.exists(archive_path):
                    logging.error(f"Archive size: {os.path.getsize(archive_path)}")
                
                # Clean up partial extraction
                if os.path.exists(frik_mod_dir):
                    try:
                        shutil.rmtree(frik_mod_dir)
                        logging.info(f"Removed partial extraction: {frik_mod_dir}")
                    except Exception as cleanup_error:
                        logging.warning(f"Failed to remove partial extraction: {cleanup_error}")
                
                # Re-raise with more context
                raise Exception(f"Failed to extract FRIK archive: {str(extract_error)}")
            
            # Clean up archive after successful extraction
            if os.path.exists(archive_path):
                os.unlink(archive_path)
            
            logging.info(f"Extracted FRIK to {frik_mod_dir}")
            self.root.after(0, lambda: self.message_label.config(text="FRIK installed successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
        except Exception as e:
            logging.error(f"Failed to extract FRIK: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to extract FRIK: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def download_and_install_comfort_swim(self):
        """Download and install Comfort Swim VR mod"""
        try:
            comfort_swim_url = "https://github.com/ArthurHub/F4VRComfortSwim/releases/download/v0.3.0/Comfort.Swim.VR.-.v0.3.0.-.20250711.7z"
            temp_dir = os.path.join(tempfile.gettempdir())
            comfort_swim_archive = os.path.join(temp_dir, "Comfort.Swim.VR.-.v0.3.0.-.20250711.7z")
            
            self.root.after(0, lambda: self.message_label.config(text="Setting up Comfort Swim VR", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Download Comfort Swim VR using the retry-enabled download function
            self.download_with_memory_management(comfort_swim_url, comfort_swim_archive, "Comfort Swim VR")
            logging.info(f"Downloaded Comfort Swim VR archive to {comfort_swim_archive}")
            
            # Extract Comfort Swim VR
            self.extract_comfort_swim(comfort_swim_archive)
            
        except Exception as e:
            logging.error(f"Comfort Swim VR installation failed: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to install Comfort Swim VR: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def extract_comfort_swim(self, archive_path):
        """Extract Comfort Swim VR to mods directory"""
        try:
            # Create Comfort Swim VR mod directory
            comfort_swim_mod_dir = os.path.join(self.mo2_path.get(), "mods", "Comfort Swim VR")
            os.makedirs(comfort_swim_mod_dir, exist_ok=True)
            
            self.root.after(0, lambda: self.message_label.config(text="Extracting Comfort Swim VR", fg="#ffffff") if self.message_label.winfo_exists() else None)
            self.create_progress_bar("Extracting Comfort Swim VR")
            
            # Use bundled 7za.exe for extraction
            bundled_7za = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "7za.exe")
            
            if not os.path.exists(bundled_7za):
                raise FileNotFoundError(f"Bundled 7za.exe not found at {bundled_7za}")
            
            # Extract using bundled 7za
            self.root.after(0, lambda: self.progress_label.config(text="Extracting Comfort Swim VR"))
            extract_cmd = [bundled_7za, "x", archive_path, f"-o{comfort_swim_mod_dir}", "-y", "-bb0", "-bd"]
            
            result = subprocess.run(extract_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode != 0:
                raise Exception(f"7za extraction failed: {result.stderr}")
            
            # Update progress to 100%
            self.root.after(0, lambda: self.progress.__setitem__("value", 100))
            self.root.after(0, lambda: self.progress_label.config(text="Extracting Comfort Swim VR (100%)"))
            
            # Clean up archive
            if os.path.exists(archive_path):
                os.unlink(archive_path)
            
            logging.info(f"Extracted Comfort Swim VR to {comfort_swim_mod_dir}")
            self.root.after(0, lambda: self.message_label.config(text="Comfort Swim VR installed successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
        except Exception as e:
            logging.error(f"Failed to extract Comfort Swim VR: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to extract Comfort Swim VR: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def download_and_install_buffout4(self):
        """Download and install Buffout 4 NG mod"""
        try:
            buffout4_url = "https://github.com/alandtse/Buffout4/releases/download/v1.37.0/Buffout4_NG-1.37.0.7z"
            temp_dir = os.path.join(tempfile.gettempdir())
            buffout4_archive = os.path.join(temp_dir, "Buffout4_NG-1.37.0.7z")
            
            self.root.after(0, lambda: self.message_label.config(text="Setting up Buffout 4 NG", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
            # Download Buffout 4 NG using the retry-enabled download function
            self.download_with_memory_management(buffout4_url, buffout4_archive, "Buffout 4 NG")
            logging.info(f"Downloaded Buffout 4 NG archive to {buffout4_archive}")
            
            # Extract Buffout 4 NG
            self.extract_buffout4(buffout4_archive)
            
        except Exception as e:
            logging.error(f"Buffout 4 NG installation failed: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to install Buffout 4 NG: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def extract_buffout4(self, archive_path):
        """Extract Buffout 4 NG to mods directory"""
        try:
            # Create Buffout 4 NG mod directory
            buffout4_mod_dir = os.path.join(self.mo2_path.get(), "mods", "Buffout 4 NG")
            os.makedirs(buffout4_mod_dir, exist_ok=True)
            
            self.root.after(0, lambda: self.message_label.config(text="Extracting Buffout 4 NG", fg="#ffffff") if self.message_label.winfo_exists() else None)
            self.create_progress_bar("Extracting Buffout 4 NG")
            
            # Use bundled 7za.exe for extraction
            bundled_7za = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "7za.exe")
            
            if not os.path.exists(bundled_7za):
                raise FileNotFoundError(f"Bundled 7za.exe not found at {bundled_7za}")
            
            # Extract using bundled 7za
            self.root.after(0, lambda: self.progress_label.config(text="Extracting Buffout 4 NG"))
            extract_cmd = [bundled_7za, "x", archive_path, f"-o{buffout4_mod_dir}", "-y", "-bb0", "-bd"]
            
            result = subprocess.run(extract_cmd, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode != 0:
                raise Exception(f"7za extraction failed: {result.stderr}")
            
            # Update progress to 100%
            self.root.after(0, lambda: self.progress.__setitem__("value", 100))
            self.root.after(0, lambda: self.progress_label.config(text="Extracting Buffout 4 NG (100%)"))
            
            # Clean up archive
            if os.path.exists(archive_path):
                os.unlink(archive_path)
            
            logging.info(f"Extracted Buffout 4 NG to {buffout4_mod_dir}")
            self.root.after(0, lambda: self.message_label.config(text="Buffout 4 NG installed successfully.", fg="#ffffff") if self.message_label.winfo_exists() else None)
            
        except Exception as e:
            logging.error(f"Failed to extract Buffout 4 NG: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to extract Buffout 4 NG: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)
            raise

    def create_desktop_shortcut(self):
        """Create a desktop shortcut to launch MO2 and start the game with custom icon"""
        try:
            import win32com.client  # Already available via your imports
            
            mo2_exe = os.path.join(self.mo2_path.get(), "ModOrganizer.exe")
            if not os.path.exists(mo2_exe):
                logging.error(f"Cannot create shortcut: ModOrganizer.exe not found at {mo2_exe}")
                return
            
            # Get desktop path
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop, "Fallout London VR.lnk")
            
            # Get icon path (from your assets)
            icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "icon.ico")
            if not os.path.exists(icon_path):
                logging.warning(f"Icon not found at {icon_path}; shortcut will use default icon")
                icon_path = mo2_exe  # Fallback to MO2's icon
            
            # Create shortcut
            shell = win32com.client.Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = mo2_exe
            shortcut.Arguments = '"moshortcut://Portable:Fallout 4 VR"'
            shortcut.WorkingDirectory = self.mo2_path.get()
            shortcut.IconLocation = f"{icon_path},0"  # 0 is the icon index
            shortcut.Description = "Launch Fallout: London VR via ModOrganizer2"
            shortcut.save()
            
            logging.info(f"Created desktop shortcut at {shortcut_path} with icon {icon_path}")
            self.root.after(0, lambda: self.message_label.config(text="") if self.message_label.winfo_exists() else None)
        
        except Exception as e:
            logging.error(f"Failed to create desktop shortcut: {e}")
            self.root.after(0, lambda es=str(e): self.message_label.config(text=f"Failed to create shortcut: {es}", fg="#ff6666") if self.message_label.winfo_exists() else None)

    def create_start_menu_shortcuts(self):
        """Create Start Menu folder with game and MO2 shortcuts"""
        try:
            import win32com.client
            
            mo2_exe = os.path.join(self.mo2_path.get(), "ModOrganizer.exe")
            if not os.path.exists(mo2_exe):
                logging.error(f"Cannot create Start Menu shortcuts: ModOrganizer.exe not found at {mo2_exe}")
                return
            
            # Get Start Menu programs folder
            start_menu = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs")
            fallout_london_folder = os.path.join(start_menu, "Fallout London VR")
            
            # Create the folder
            os.makedirs(fallout_london_folder, exist_ok=True)
            logging.info(f"Created Start Menu folder: {fallout_london_folder}")
            
            # Get icon path
            icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "icon.ico")
            if not os.path.exists(icon_path):
                logging.warning(f"Icon not found at {icon_path}; shortcuts will use default icons")
                icon_path = mo2_exe  # Fallback to MO2's icon
            
            shell = win32com.client.Dispatch('WScript.Shell')
            
            # Create "Fallout London VR" shortcut (launches game)
            game_shortcut_path = os.path.join(fallout_london_folder, "Fallout London VR.lnk")
            game_shortcut = shell.CreateShortCut(game_shortcut_path)
            game_shortcut.Targetpath = mo2_exe
            game_shortcut.Arguments = '"moshortcut://Portable:Fallout 4 VR"'
            game_shortcut.WorkingDirectory = self.mo2_path.get()
            game_shortcut.IconLocation = f"{icon_path},0"
            game_shortcut.Description = "Launch Fallout: London VR via ModOrganizer2"
            game_shortcut.save()
            logging.info(f"Created game shortcut: {game_shortcut_path}")
            
            # Create "Mod Organizer 2" shortcut (launches MO2 directly)
            mo2_shortcut_path = os.path.join(fallout_london_folder, "Mod Organizer 2.lnk")
            mo2_shortcut = shell.CreateShortCut(mo2_shortcut_path)
            mo2_shortcut.Targetpath = mo2_exe
            mo2_shortcut.Arguments = ""  # No arguments - just open MO2
            mo2_shortcut.WorkingDirectory = self.mo2_path.get()
            mo2_shortcut.IconLocation = f"{mo2_exe},0"  # Use MO2's own icon
            mo2_shortcut.Description = "Open Mod Organizer 2 for Fallout: London VR"
            mo2_shortcut.save()
            logging.info(f"Created MO2 shortcut: {mo2_shortcut_path}")

            # Create "Donation page" shortcut that opens the donation URL in the user's default browser
            try:
                donation_shortcut_path = os.path.join(fallout_london_folder, "Donation page.lnk")
                donation_url = "https://ko-fi.com/falloutlondonvr"

                # We'll create a small .lnk that points to rundll32 url.dll,FileProtocolHandler <url>
                # This approach opens the URL using the default browser.
                donation_shortcut = shell.CreateShortCut(donation_shortcut_path)
                donation_shortcut.Targetpath = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'System32', 'rundll32.exe')
                donation_shortcut.Arguments = f"url.dll,FileProtocolHandler {donation_url}"
                donation_shortcut.WorkingDirectory = os.path.expanduser("~")
                # Use the icon from the assets if available, else let Windows pick default browser icon
                favicon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), "assets", "icon.ico")
                if os.path.exists(favicon_path):
                    donation_shortcut.IconLocation = f"{favicon_path},0"
                else:
                    donation_shortcut.IconLocation = f"{mo2_exe},0"
                donation_shortcut.Description = "Open the Fallout London VR donation page"
                donation_shortcut.save()
                logging.info(f"Created donation shortcut: {donation_shortcut_path}")
            except Exception as e:
                logging.warning(f"Failed to create donation Start Menu shortcut: {e}")
            
            logging.info(f"Successfully created Start Menu shortcuts in {fallout_london_folder}")
            
        except Exception as e:
            logging.error(f"Failed to create Start Menu shortcuts: {e}")
            # Non-critical error, don't show to user


def main():
    """Main function to start the installer"""
    try:
        # Check for --fresh-install flag
        skip_update = '--fresh-install' in sys.argv
        if skip_update:
            logging.info("Starting in fresh install mode (--fresh-install flag detected)")
        
        # Check for --install-path argument
        initial_path = None
        if '--install-path' in sys.argv:
            try:
                path_index = sys.argv.index('--install-path') + 1
                if path_index < len(sys.argv):
                    initial_path = sys.argv[path_index]
                    logging.info(f"Using preserved install path: {initial_path}")
            except (ValueError, IndexError):
                pass
        
        root = tk.Tk()
        app = FalloutLondonVRInstaller(root, skip_update_detection=skip_update, initial_install_path=initial_path)
        root.mainloop()
    except PermissionError as e:
        # Handle SSL/certificate permission errors specifically
        if "virtual_file.log" in str(e) or "Volume{" in str(e):
            error_msg = "Permission error. Please run installer again as an admin."
        else:
            error_msg = f"Permission error: {e}\n\nPlease run installer again as an admin."
        
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Permission Error", error_msg)
        except:
            # Fallback if tkinter fails
            print(f"ERROR: {error_msg}")
        
        logging.error(f"Permission error in main: {e}")
    except Exception as e:
        logging.error(f"Fatal error in main: {e}")
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Fatal Error", f"A fatal error occurred: {e}")
        except:
            # Fallback if tkinter fails
            print(f"FATAL ERROR: {e}")

if __name__ == "__main__":
    main()