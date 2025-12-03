    def copy_directory_with_robocopy(self, src_dir, dest_dir, label_text, exclude_dirs=None):
        """Fast directory copy using robocopy with progress tracking"""
        if exclude_dirs is None:
            exclude_dirs = ["F4SE", "source", "scripts\\source"]
        
        self.create_progress_bar(label_text)
        
        try:
            # Build robocopy command
            cmd = [
                'robocopy',
                src_dir,
                dest_dir,
                '/E',         # Copy subdirectories, including empty ones
                '/Z',         # Restartable mode (better for large files)
                '/MT:16',     # Use 16 threads (significant speed boost)
                '/R:2',       # Retry 2 times
                '/W:1',       # Wait 1 second between retries
                '/NFL',       # No file list (less output to parse)
                '/NDL',       # No directory list
                '/NJH',       # No job header
                '/NJS',       # No job summary
                '/NP',        # No progress percentage in output
                '/BYTES'      # Show sizes in bytes
            ]
            
            # Add exclusions
            if exclude_dirs:
                cmd.extend(['/XD'] + exclude_dirs)
            
            # Count total size first for progress bar
            total_size = self.get_directory_size(src_dir, exclude_dirs)
            copied_size = 0
            
            # Run robocopy with real-time output parsing
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )
            
            # Parse output for progress
            for line in process.stdout:
                if self.cancel_requested:
                    process.terminate()
                    raise InterruptedError("Installation cancelled by user")
                
                # Parse robocopy output for file sizes
                # Format: "        1234  filename.ext"
                parts = line.strip().split(None, 1)
                if len(parts) == 2 and parts[0].isdigit():
                    file_size = int(parts[0])
                    copied_size += file_size
                    
                    # Update progress
                    if total_size > 0:
                        progress = min((copied_size / total_size) * 100, 100)
                        self.root.after(0, lambda p=progress: self.progress.__setitem__("value", p))
                        self.root.after(0, lambda p=progress: self.progress_label.config(
                            text=f"{label_text} ({p:.1f}%)"
                        ))
            
            # Wait for process to complete
            return_code = process.wait()
            
            # Robocopy return codes:
            # 0 = No files copied
            # 1 = Files copied successfully
            # 2 = Extra files/directories detected
            # 4 = Mismatched files/directories detected
            # 8 = Some files/directories could not be copied
            # 16 = Fatal error
            
            if return_code >= 8:
                raise Exception(f"Robocopy failed with code {return_code}")
            
            logging.info(f"Robocopy completed successfully with code {return_code}")
            
        except Exception as e:
            logging.error(f"Robocopy failed: {e}")
            # Fallback to Python copy
            logging.info("Falling back to Python copy method...")
            self.copy_directory_with_progress_python(src_dir, dest_dir, label_text, exclude_dirs)

    def copy_directory_with_progress_python(self, src_dir, dest_dir, label_text, exclude_dirs=None):
        """Original Python-based copy method as fallback"""
        # Call the existing copy_directory_with_progress method
        self.copy_directory_with_progress(src_dir, dest_dir, label_text, exclude_dirs)

    def get_directory_size(self, path, exclude_dirs=None):
        """Calculate total size of directory for progress tracking"""
        if exclude_dirs is None:
            exclude_dirs = []
        
        total_size = 0
        exclude_dirs_lower = [d.lower() for d in exclude_dirs]
        
        for dirpath, dirnames, filenames in os.walk(path):
            # Remove excluded directories
            dirnames[:] = [d for d in dirnames if d.lower() not in exclude_dirs_lower]
            
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                try:
                    total_size += os.path.getsize(filepath)
                except OSError:
                    pass
        
        return total_size

    def copy_fallout_data_fast(self):
        """Copy Fallout data files using robocopy for speed"""
        try:
            self.root.after(0, lambda: self.message_label.config(
                text="Preparing Game Assets (Fast Mode)", fg="#ffffff"
            ) if self.message_label.winfo_exists() else None)
            
            logging.info("Starting fast Fallout data copy using robocopy")
            
            src_f4_data = os.path.join(self.f4_path.get(), "Data")
            src_london_data = None
            if self.london_data_path.get() and self.london_data_path.get() != "Already installed":
                src_london_data = os.path.join(self.london_data_path.get(), "Data")
            
            # Create mods directory
            mods_dir = os.path.join(self.mo2_path.get(), "mods")
            os.makedirs(mods_dir, exist_ok=True)
            
            # Scenario 1: F4 dir + separate London dir
            if not self.london_installed and src_london_data and os.path.exists(src_london_data):
                f4_mod_dir = os.path.join(mods_dir, "Fallout 4 Data")
                london_mod_dir = os.path.join(mods_dir, "Fallout London Data")
                
                # Copy Fallout 4 Data
                if os.path.exists(src_f4_data):
                    self.copy_directory_with_robocopy(
                        src_f4_data, f4_mod_dir, "Copying Fallout 4 Data (Fast)"
                    )
                
                # Copy Fallout London Data
                if os.path.exists(src_london_data):
                    self.copy_directory_with_robocopy(
                        src_london_data, london_mod_dir, "Copying Fallout: London Data (Fast)"
                    )
            else:
                # Scenario 2: Combined F4+London data
                london_mod_dir = os.path.join(mods_dir, "Fallout London Data")
                if os.path.exists(src_f4_data):
                    self.copy_directory_with_robocopy(
                        src_f4_data, london_mod_dir, "Copying Fallout: London Data (Fast)"
                    )
            
            self.root.after(0, lambda: self.message_label.config(
                text="Fallout data files copied successfully.", fg="#ffffff"
            ) if self.message_label.winfo_exists() else None)
            
            logging.info("Fast Fallout data copy completed")
            
        except Exception as e:
            self.root.after(0, lambda es=str(e): self.message_label.config(
                text=f"Failed to copy Fallout data: {es}", fg="#ff6666"
            ) if self.message_label.winfo_exists() else None)
            logging.error(f"Fast Fallout data copy failed: {e}")
            raise
