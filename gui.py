#!/usr/bin/env python3
"""
PASS Schedule PDF Generator - GUI Interface
Simple graphical interface with logs, progress bar and start button.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import sys
import os
from datetime import datetime
import subprocess
from dotenv import load_dotenv
import configparser
import base64
from PIL import Image, ImageTk


class PDFGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PASS Schedule PDF Generator")
        self.root.geometry("800x900")  # Increased height to accommodate all settings and progress
        self.root.resizable(True, True)
        
        # Communication queue between threads
        self.log_queue = queue.Queue()
        
        # Progress tracking
        self.is_running = False
        
        # Settings file path
        self.settings_file = "settings.ini"
        
        # Flag to prevent auto-save indication during loading
        self.loading_settings = False
        
        # Track original password to detect real changes
        self.original_password = ""
        
        # Setup UI
        self.setup_ui()
        
        # Load settings
        self.load_settings()
        
        # Start queue processing
        self.process_queue()
        
        # Load environment variables
        load_dotenv()
        
    def setup_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)  # Progress frame gets the extra space
        
        # Header frame for title and logo
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(
            header_frame, 
            text="PASS Schedule PDF Generator", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Logo
        try:
            logo_path = os.path.join("resources", "fiplogopixel.png")
            if os.path.exists(logo_path):
                # Load and resize logo
                logo_image = Image.open(logo_path)
                # Resize to reasonable size (max 80px height)
                logo_height = 80
                logo_width = int(logo_image.width * logo_height / logo_image.height)
                logo_image = logo_image.resize((logo_width, logo_height), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                
                logo_label = ttk.Label(header_frame, image=self.logo_photo)
                logo_label.grid(row=0, column=1, sticky=tk.E, padx=(10, 0))
        except Exception as e:
            print(f"Could not load logo: {e}")
            # Fallback text logo
            logo_label = ttk.Label(
                header_frame, 
                text="FIPA", 
                font=("Arial", 12, "bold"), 
                foreground="#0066cc"
            )
            logo_label.grid(row=0, column=1, sticky=tk.E, padx=(10, 0))
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Paramètres", padding="10")
        settings_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)
        
        # PASS Username field
        ttk.Label(settings_frame, text="Nom d'utilisateur PASS:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.username_var = tk.StringVar()
        self.username_var.trace('w', self.on_setting_changed)
        self.username_entry = ttk.Entry(settings_frame, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # PASS Password field
        ttk.Label(settings_frame, text="Mot de passe PASS:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.password_var = tk.StringVar()
        self.password_var.trace('w', self.on_password_changed)  # Special handler for password
        self.password_entry = ttk.Entry(settings_frame, textvariable=self.password_var, show="*", width=30)
        self.password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # Save password checkbox
        self.save_password_var = tk.BooleanVar()
        self.save_password_cb = ttk.Checkbutton(
            settings_frame, 
            text="Sauvegarder le mot de passe (en clair, non sécurisé)",
            variable=self.save_password_var,
            command=self.on_save_password_changed
        )
        self.save_password_cb.grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        
        # Separator
        ttk.Separator(settings_frame, orient='horizontal').grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        
        # Personal information section
        ttk.Label(settings_frame, text="Nom et Prénom:").grid(row=4, column=0, sticky=tk.W, padx=(0, 10))
        self.nom_prenom_var = tk.StringVar()
        self.nom_prenom_var.trace('w', self.on_setting_changed)
        self.nom_prenom_entry = ttk.Entry(settings_frame, textvariable=self.nom_prenom_var, width=30)
        self.nom_prenom_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(settings_frame, text="Promotion:").grid(row=5, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.promo_var = tk.StringVar()
        self.promo_var.trace('w', self.on_setting_changed)
        self.promo_entry = ttk.Entry(settings_frame, textvariable=self.promo_var, width=30)
        self.promo_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        ttk.Label(settings_frame, text="Semaine cible:").grid(row=6, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.target_week_var = tk.StringVar()
        self.target_week_var.trace('w', self.on_setting_changed)
        self.target_week_spinbox = ttk.Spinbox(
            settings_frame, 
            textvariable=self.target_week_var, 
            from_=1, 
            to=53,
            width=10,
            justify='center'
        )
        self.target_week_spinbox.grid(row=6, column=1, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        
        # Add current week info
        current_week = datetime.now().isocalendar()[1]
        week_info_label = ttk.Label(
            settings_frame, 
            text=f"(Semaine actuelle: {current_week})",
            font=("Segoe UI", 8),
            foreground="gray"
        )
        week_info_label.grid(row=6, column=1, sticky=tk.W, padx=(120, 0), pady=(5, 0))
        
        # Separator
        ttk.Separator(settings_frame, orient='horizontal').grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 10))
        
        # PDF settings section
        ttk.Label(settings_frame, text="Message PDF:").grid(row=8, column=0, sticky=tk.W, padx=(0, 10))
        self.pdf_message_var = tk.StringVar()
        self.pdf_message_var.trace('w', self.on_setting_changed)
        self.pdf_message_entry = ttk.Entry(settings_frame, textvariable=self.pdf_message_var, width=30)
        self.pdf_message_entry.grid(row=8, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(settings_frame, text="Fichier signature:").grid(row=9, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.signature_file_var = tk.StringVar()
        self.signature_file_var.trace('w', self.on_setting_changed)
        signature_frame = ttk.Frame(settings_frame)
        signature_frame.grid(row=9, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        signature_frame.columnconfigure(0, weight=1)
        
        self.signature_file_entry = ttk.Entry(signature_frame, textvariable=self.signature_file_var)
        self.signature_file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.browse_signature_btn = ttk.Button(signature_frame, text="Parcourir", command=self.browse_signature_file)
        self.browse_signature_btn.grid(row=0, column=1)
        
        # Save settings button
        buttons_frame = ttk.Frame(settings_frame)
        buttons_frame.grid(row=10, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        buttons_frame.columnconfigure(1, weight=1)
        
        self.reset_btn = ttk.Button(
            buttons_frame, 
            text="🔄 Réinitialiser", 
            command=self.reset_to_defaults
        )
        self.reset_btn.grid(row=0, column=0, padx=(0, 5))
        
        self.save_settings_btn = ttk.Button(
            buttons_frame, 
            text="💾 Sauvegarder", 
            command=self.save_settings_manually,
            style="Accent.TButton",
            state="disabled"  # Start disabled
        )
        self.save_settings_btn.grid(row=0, column=1, sticky=tk.E)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(2, weight=1)  # Log area gets the extra space
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            length=400
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Current step label
        self.step_var = tk.StringVar(value="Prêt à générer le PDF")
        self.step_label = ttk.Label(
            progress_frame,
            textvariable=self.step_var,
            font=("Segoe UI", 8),
            foreground="#2563eb"
        )
        self.step_label.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(1, 5))
        
        # Log text area with scrollbar
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=20,
            width=80,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Control buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(1, weight=1)
        
        # Start button
        self.start_button = ttk.Button(
            button_frame,
            text="Générer PDF",
            command=self.start_generation,
            style="Accent.TButton"
        )
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        # Clear logs button
        self.clear_button = ttk.Button(
            button_frame,
            text="Effacer logs",
            command=self.clear_logs
        )
        self.clear_button.grid(row=0, column=1, padx=(0, 10))
        
        # Open folder button
        self.open_folder_button = ttk.Button(
            button_frame,
            text="Ouvrir dossier",
            command=self.open_output_folder
        )
        self.open_folder_button.grid(row=0, column=2)
        
    
        
        # Initial log message
        self.log_message("🎯 PASS Schedule PDF Generator ready!")
        self.log_message("📋 Configurez vos paramètres ci-dessus")
        self.log_message("▶️ Click 'Generate PDF' to start")
        
    def browse_signature_file(self):
        """Open file dialog to select signature file"""
        from tkinter import filedialog
        
        filename = filedialog.askopenfilename(
            title="Sélectionner le fichier de signature",
            filetypes=[
                ("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.signature_file_var.set(filename)
            self.save_settings()
    
    def on_password_changed(self, *args):
        """Called when password field changes - only triggers save indication if save password is enabled"""
        # Ignore changes during settings loading
        if self.loading_settings:
            return
        
        # Only trigger save indication if save password is checked
        if self.save_password_var.get():
            # Check if password actually changed from original
            current_password = self.password_var.get()
            if current_password != self.original_password:
                # Change save button text to indicate unsaved changes
                if hasattr(self, 'save_settings_btn'):
                    self.save_settings_btn.config(text="💾 Sauvegarder*", state="normal")
    
    def on_setting_changed(self, *args):
        """Called when any setting field changes - enables auto-save"""
        # Ignore changes during settings loading
        if self.loading_settings:
            return
            
        # Change save button text to indicate unsaved changes
        if hasattr(self, 'save_settings_btn'):
            self.save_settings_btn.config(text="💾 Sauvegarder*", state="normal")
        
    def load_settings(self):
        """Load settings from INI file"""
        self.loading_settings = True  # Disable change tracking during loading
        config = configparser.ConfigParser()
        
        if os.path.exists(self.settings_file):
            try:
                config.read(self.settings_file, encoding='utf-8')
                
                # Load PASS credentials
                if config.has_option('PASS', 'username'):
                    self.username_var.set(config.get('PASS', 'username'))
                
                # Load password if saved
                if config.has_option('PASS', 'save_password'):
                    save_password = config.getboolean('PASS', 'save_password')
                    self.save_password_var.set(save_password)
                    
                    if save_password and config.has_option('PASS', 'password'):
                        # Decode password (simple base64 encoding)
                        encoded_password = config.get('PASS', 'password')
                        try:
                            password = base64.b64decode(encoded_password.encode()).decode()
                            self.password_var.set(password)
                            self.original_password = password  # Track original password
                        except Exception as e:
                            self.log_message(f"⚠️ Erreur lors du décodage du mot de passe: {e}")
                
                # Load personal information
                if config.has_option('PERSONAL', 'nom_prenom'):
                    self.nom_prenom_var.set(config.get('PERSONAL', 'nom_prenom'))
                if config.has_option('PERSONAL', 'promo'):
                    self.promo_var.set(config.get('PERSONAL', 'promo'))
                if config.has_option('PERSONAL', 'target_week'):
                    self.target_week_var.set(config.get('PERSONAL', 'target_week'))
                
                # Load PDF settings
                if config.has_option('PDF', 'message'):
                    self.pdf_message_var.set(config.get('PDF', 'message'))
                if config.has_option('PDF', 'signature_file'):
                    self.signature_file_var.set(config.get('PDF', 'signature_file'))
                
                self.log_message("✅ Paramètres chargés depuis settings.ini")
                
            except Exception as e:
                self.log_message(f"⚠️ Erreur lors du chargement des paramètres: {e}")
        else:
            # Load defaults from .env file if settings.ini doesn't exist
            self.load_defaults_from_env()
            self.log_message("📝 Fichier settings.ini non trouvé - chargement des valeurs par défaut depuis .env")
            # Auto-save the defaults to create the ini file
            self.save_settings()
        
        # Re-enable change tracking after loading is complete
        self.loading_settings = False
        
        # Ensure save button starts disabled after loading
        if hasattr(self, 'save_settings_btn'):
            self.save_settings_btn.config(text="💾 Sauvegarder", state="disabled")
    
    def load_defaults_from_env(self):
        """Load default values from .env file"""
        # Load username/password from .env
        username = os.getenv('IMT_USERNAME', '')
        password = os.getenv('IMT_PASSWORD', '')
        
        if username:
            self.username_var.set(username)
        if password:
            self.password_var.set(password)
            self.original_password = password  # Track original password
            
        # Load other settings with sensible defaults
        self.nom_prenom_var.set(os.getenv('NOM_PRENOM', 'VOTRENOM Prénom'))
        self.promo_var.set(os.getenv('PROMO', 'FIPA3R'))
        self.target_week_var.set(os.getenv('TARGET_WEEK', '38'))
        self.pdf_message_var.set(os.getenv('PDF_MESSAGE', 'Certifie sur l\'honneur avoir été présent sur les créneaux indiqués dans le planning'))
        self.signature_file_var.set(os.getenv('SIGNATURE_FILE', 'signature.png'))
    
    def reset_to_defaults(self):
        """Reset all settings to default values from .env"""
        result = messagebox.askyesno(
            "Confirmation", 
            "Réinitialiser tous les paramètres aux valeurs par défaut ?\n\nCela remplacera tous les paramètres actuels."
        )
        if result:
            self.loading_settings = True  # Disable change tracking during reset
            self.load_defaults_from_env()
            self.original_password = self.password_var.get()  # Update original after reset
            self.loading_settings = False  # Re-enable change tracking
            # Reset save button since we just loaded defaults
            if hasattr(self, 'save_settings_btn'):
                self.save_settings_btn.config(text="💾 Sauvegarder", state="disabled")
            self.log_message("🔄 Paramètres réinitialisés aux valeurs par défaut")
    
    def save_settings_manually(self):
        """Save settings manually when user clicks save button"""
        self.save_settings()
        # Update original password after successful save
        self.original_password = self.password_var.get()
        self.save_settings_btn.config(text="💾 Sauvegarder", state="disabled")
        messagebox.showinfo("Succès", "Paramètres sauvegardés avec succès !")
    
    def save_settings(self):
        """Save settings to INI file"""
        config = configparser.ConfigParser()
        
        # Load existing settings first
        if os.path.exists(self.settings_file):
            config.read(self.settings_file, encoding='utf-8')
        
        # Ensure sections exist
        for section in ['PASS', 'PERSONAL', 'PDF']:
            if not config.has_section(section):
                config.add_section(section)
        
        # Save PASS credentials
        config.set('PASS', 'username', self.username_var.get())
        config.set('PASS', 'save_password', str(self.save_password_var.get()))
        
        # Save password if requested
        if self.save_password_var.get():
            password = self.password_var.get()
            if password:
                encoded_password = base64.b64encode(password.encode()).decode()
                config.set('PASS', 'password', encoded_password)
        else:
            # Remove password if not saving
            if config.has_option('PASS', 'password'):
                config.remove_option('PASS', 'password')
        
        # Save personal information
        config.set('PERSONAL', 'nom_prenom', self.nom_prenom_var.get())
        config.set('PERSONAL', 'promo', self.promo_var.get())
        config.set('PERSONAL', 'target_week', self.target_week_var.get())
        
        # Save PDF settings
        config.set('PDF', 'message', self.pdf_message_var.get())
        config.set('PDF', 'signature_file', self.signature_file_var.get())
        
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                config.write(f)
            self.log_message("💾 Paramètres sauvegardés dans settings.ini")
        except Exception as e:
            self.log_message(f"❌ Erreur lors de la sauvegarde: {e}")
    
    def on_save_password_changed(self):
        """Handle save password checkbox change"""
        # Don't auto-save, just manage the UI state
        if not self.save_password_var.get():
            # If unchecked, ask for confirmation to remove saved password
            result = messagebox.askyesno(
                "Confirmation", 
                "Supprimer le mot de passe sauvegardé ?\n\nVous devrez le ressaisir à la prochaine utilisation."
            )
            if not result:
                # User cancelled, recheck the box
                self.save_password_var.set(True)
                return
        
        # Check if we need to update save button status based on current password state
        if self.save_password_var.get():
            # If we're now saving passwords, check if current password differs from original
            current_password = self.password_var.get()
            if current_password != self.original_password:
                if hasattr(self, 'save_settings_btn'):
                    self.save_settings_btn.config(text="💾 Sauvegarder*", state="normal")
        
        # Just mark as needing save for the checkbox change itself
        if hasattr(self, 'save_settings_btn'):
            self.save_settings_btn.config(text="💾 Sauvegarder*", state="normal")
        
    def log_message(self, message):
        """Add a message to the log area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def update_progress(self, value, step=""):
        """Update progress bar and current step"""
        self.progress_var.set(value)
        if step:
            self.step_var.set(step)
        self.root.update_idletasks()
        
    def start_generation(self):
        """Start PDF generation in a separate thread"""
        if self.is_running:
            self.log_message("⚠️ Generation already in progress!")
            return
        
        # Validate required settings
        if not self.username_var.get().strip():
            messagebox.showerror("Erreur", "Veuillez saisir votre nom d'utilisateur PASS")
            return
            
        if not self.password_var.get().strip():
            messagebox.showerror("Erreur", "Veuillez saisir votre mot de passe PASS")
            return
            
        if not self.nom_prenom_var.get().strip():
            messagebox.showerror("Erreur", "Veuillez saisir votre nom et prénom")
            return
            
        if not self.promo_var.get().strip():
            messagebox.showerror("Erreur", "Veuillez saisir votre promotion")
            return
            
        if not self.target_week_var.get().strip():
            messagebox.showerror("Erreur", "Veuillez saisir la semaine cible")
            return
            
        # Validate week number
        try:
            week_num = int(self.target_week_var.get())
            if week_num < 1 or week_num > 53:
                messagebox.showerror("Erreur", "La semaine doit être entre 1 et 53")
                return
        except ValueError:
            messagebox.showerror("Erreur", "La semaine doit être un nombre entier")
            return
        
        # Save settings
        self.save_settings()
            
        # Disable start button
        self.start_button.config(state="disabled", text="Génération...")
        self.is_running = True
        
        # Reset progress
        self.update_progress(0, "Initialisation")
        
        # Start generation in separate thread
        thread = threading.Thread(target=self.generate_pdf_thread, daemon=True)
        thread.start()
        
    def generate_pdf_thread(self):
        """Run PDF generation by calling the script as subprocess"""
        try:
            self.log_queue.put(('progress', 5, "Préparation de l'environnement"))
            self.log_queue.put(('log', "🚀 Starting PDF generation process...\n"))
            print("🚀 Starting PDF generation from GUI...")  # Console log
            
            # Run the script as subprocess to avoid import issues
            import subprocess
            import sys
            
            # Get Python executable path (use current Python interpreter)
            import sys
            python_exe = sys.executable
            
            # Si on détecte qu'on est dans un venv, utilisons l'exécutable du venv
            if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
                # Nous sommes dans un environnement virtuel
                print(f"📦 Virtual environment detected: {python_exe}")
            else:
                # Essayer de trouver le venv local
                venv_python = os.path.join(os.getcwd(), ".venv", "Scripts", "python.exe")
                if os.path.exists(venv_python):
                    python_exe = venv_python
                    print(f"📦 Using local virtual environment: {python_exe}")
                else:
                    print(f"🐍 Using system Python: {python_exe}")
                    print("⚠️  Warning: Make sure all dependencies are installed in system Python")
                
            script_path = "pass-schedule-pdf.py"
            
            self.log_queue.put(('progress', 15, "Lancement du script"))
            print(f"📄 Executing: {python_exe} {script_path}")  # Console log
            
            # Prepare environment with UTF-8 encoding for emojis
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONLEGACYWINDOWSSTDIO'] = '0'  # Force UTF-8 on Windows
            
            # Override all settings with GUI values
            env['IMT_USERNAME'] = self.username_var.get()
            env['IMT_PASSWORD'] = self.password_var.get()
            env['NOM_PRENOM'] = self.nom_prenom_var.get()
            env['PROMO'] = self.promo_var.get()
            env['TARGET_WEEK'] = self.target_week_var.get()
            env['PDF_MESSAGE'] = self.pdf_message_var.get()
            env['SIGNATURE_FILE'] = self.signature_file_var.get()
            # Keep DEBUG_MODE from .env or default to false
            env['DEBUG_MODE'] = os.getenv('DEBUG_MODE', 'false')
            
            # Start the subprocess with separate stdout/stderr
            process = subprocess.Popen(
                [python_exe, script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',  # Replace problematic characters instead of crashing
                cwd=os.getcwd(),
                env=env
            )
            
            # Read output line by line from both stdout and stderr
            def read_stdout():
                while True:
                    line = process.stdout.readline()
                    if line == '' and process.poll() is not None:
                        break
                    if line:
                        self.log_queue.put(('log', f"[STDOUT] {line}"))
                        print(f"[STDOUT] {line.rstrip()}")
                        # Détecter les étapes dans les logs
                        self.detect_step_from_log(line)
            
            def read_stderr():
                while True:
                    line = process.stderr.readline()
                    if line == '' and process.poll() is not None:
                        break
                    if line:
                        self.log_queue.put(('log', f"[STDERR] {line}"))
                        print(f"[STDERR] {line.rstrip()}")
                        # Détecter les étapes dans les logs d'erreur aussi
                        self.detect_step_from_log(line)
            
            # Start threads to read both streams
            stdout_thread = threading.Thread(target=read_stdout, daemon=True)
            stderr_thread = threading.Thread(target=read_stderr, daemon=True)
            stdout_thread.start()
            stderr_thread.start()
            
            # Wait for process to complete
            return_code = process.wait()
            
            # Wait a bit for threads to finish reading
            stdout_thread.join(timeout=1)
            stderr_thread.join(timeout=1)
            
            if return_code == 0:
                self.log_queue.put(('progress', 100, "Génération terminée avec succès"))
                self.log_queue.put(('log', "Tâche terminée, vérifiez les logs pour plus de détails.\n"))
                print("✅ PDF generation completed successfully!")  # Console log
            else:
                self.log_queue.put(('progress', 0, "Échec de la génération"))
                self.log_queue.put(('log', "❌ Génération PDF échouée - vérifiez les logs\n"))
                print(f"❌ PDF generation failed with return code: {return_code}")  # Console log
                
        except Exception as e:
            self.log_queue.put(('log', f"❌ Error during generation: {e}\n"))
            self.log_queue.put(('progress', 0, "Erreur lors de la génération"))
        finally:
            # Re-enable button
            self.log_queue.put(('button_enable', None))
            self.is_running = False
    
    def detect_step_from_log(self, log_line):
        """Detect current step from log output and update progress"""
        log_lower = log_line.lower()
        
        # Définir les étapes et leurs patterns
        step_patterns = [
            ("starting schedule pdf generation", 20, "Démarrage du processus"),
            ("configuration validated", 25, "Configuration validée"),
            ("chrome browser started", 30, "Navigateur lancé"),
            ("connecting to pass", 35, "Connexion à PASS"),
            ("login successful", 45, "Connexion réussie"),
            ("navigating to schedule", 50, "Accès à l'emploi du temps"),
            ("navigation successful", 60, "Navigation terminée"),
            ("generating pdf", 70, "Génération du PDF en cours"),
            ("pdf generated successfully", 85, "PDF généré avec succès"),
            ("pdf de l'emploi du temps généré avec succès", 95, "Génération terminée"),
            ("schedule pdf generation completed", 100, "Processus terminé")
        ]
        
        for pattern, progress, step_text in step_patterns:
            if pattern in log_lower:
                self.log_queue.put(('progress', progress, step_text))
                break
            
    def clear_logs(self):
        """Clear the log text area"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("🗑️ Logs cleared")
        
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        save_folder = os.getenv('SAVE_FOLDER', 'pdfs')
        folder_path = os.path.abspath(save_folder)
        
        if os.path.exists(folder_path):
            # Open folder in Windows Explorer
            if os.name == 'nt':  # Windows
                os.startfile(folder_path)
            else:
                # For other OS
                subprocess.run(['xdg-open' if os.name == 'posix' else 'open', folder_path])
            self.log_message(f"📁 Opened folder: {folder_path}")
        else:
            self.log_message(f"⚠️ Folder does not exist: {folder_path}")
            
    def process_queue(self):
        """Process messages from the background thread"""
        try:
            while True:
                try:
                    item = self.log_queue.get_nowait()
                    
                    if item[0] == 'log':
                        # Add to log area
                        self.log_text.insert(tk.END, item[1])
                        self.log_text.see(tk.END)
                        
                    elif item[0] == 'progress':
                        # Update progress bar and step
                        if len(item) >= 3:
                            # Format: ('progress', value, step)
                            self.update_progress(item[1], item[2])
                        else:
                            # Format: ('progress', value)
                            self.update_progress(item[1])
                            
                    elif item[0] == 'button_enable':
                        self.start_button.config(state="normal", text="Générer PDF")
                        
                except queue.Empty:
                    break
                    
        except Exception as e:
            print(f"Error processing queue: {e}")
            
        # Schedule next check
        self.root.after(100, self.process_queue)


def main():
    """Main function to run the GUI"""
    try:
        # Create main window
        root = tk.Tk()
        
        # Set style for modern look
        style = ttk.Style()
        style.theme_use('vista' if 'vista' in style.theme_names() else 'clam')
        
        # Create application
        app = PDFGeneratorGUI(root)
        
        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Run the application
        root.mainloop()
        
    except KeyboardInterrupt:
        print("Application interrupted by user")
    except Exception as e:
        print(f"Error starting GUI: {e}")
        messagebox.showerror("Error", f"Failed to start application:\n{e}")


if __name__ == "__main__":
    main()