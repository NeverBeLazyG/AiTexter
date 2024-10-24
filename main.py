import customtkinter as ctk
from PIL import Image
import keyboard
import pyperclip
import json
import requests
import threading
import os
import time
from tkinter import messagebox
import sys
from pystray import MenuItem as item
import pystray
import openai
import win32clipboard
import win32con
import win32gui
import win32api

class TextProcessor:
    def __init__(self):
        self.settings_path = os.path.join(os.path.dirname(__file__), 'settings.json')
        self.settings = self.load_settings()
        self.selected_text = None
        self.original_clipboard = None
        self.root = ctk.CTk()
        self.root.withdraw()  # Hide main window initially
        self.settings_window = None
        self.setup_tray()
        self.setup_hotkey()
        self.create_main_window()
        
    def load_settings(self):
        try:
            with open(self.settings_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            # Default settings
            default_settings = {
                'use_openai': True,
                'openai_key': '',
                'ollama_url': 'http://localhost:11434',
                'ollama_model': 'mistral'
            }
            # Save default settings
            with open(self.settings_path, 'w') as f:
                json.dump(default_settings, f, indent=4)
            return default_settings
            
    def save_settings(self):
        with open(self.settings_path, 'w') as f:
            json.dump(self.settings, f, indent=4)

    def setup_tray(self):
        # Load icon from file
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
        icon_image = Image.open(icon_path)

        # Create menu
        menu = (
            item('Einstellungen', self.show_settings),
            item('About', self.show_about),
            item('Beenden', self.quit_app)
        )

        # Create icon
        self.icon = pystray.Icon("TextProcessor", icon_image, "Text Processor", menu)
        
        # Run icon in separate thread
        threading.Thread(target=self.icon.run, daemon=True).start()

    def setup_hotkey(self):
        keyboard.add_hotkey('shift+ctrl+0', self.capture_and_show, suppress=True)

    def create_main_window(self):
        self.root.title("Text Processor")
        self.root.geometry("300x400")
        
        frame = ctk.CTkFrame(self.root)
        frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        ctk.CTkLabel(frame, text="Text Processor", font=("Arial", 20)).pack(pady=10)
        
        buttons = [
            ("Umformulieren", self.rephrase),
            ("Übersetzen", self.translate),
            ("Zusammenfassen", self.summarize),
            ("Umformulieren + Übersetzen", self.rephrase_and_translate)
        ]
        
        for text, command in buttons:
            ctk.CTkButton(frame, text=text, command=command).pack(pady=10, padx=20, fill="x")

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def capture_and_show(self):
        # Save current clipboard content
        self.original_clipboard = pyperclip.paste()
        
        # Simulate Ctrl+C to copy selected text
        keyboard.send('ctrl+c')
        time.sleep(0.1)  # Wait for clipboard to update
        
        # Get the selected text from clipboard
        self.selected_text = pyperclip.paste()
        
        # If no text was selected, use clipboard content
        if not self.selected_text.strip() or (self.original_clipboard and self.selected_text == self.original_clipboard):
            self.selected_text = self.original_clipboard
            if not self.selected_text.strip():
                messagebox.showerror("Fehler", "Zwischenablage leer.")
                return
        
        # Show the window
        self.show_window()

    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def hide_window(self):
        self.root.withdraw()

    def show_settings(self, _=None):
        if self.settings_window is None:
            self.settings_window = ctk.CTkToplevel(self.root)
            self.settings_window.title("Einstellungen")
            self.settings_window.geometry("400x500")
            self.settings_window.protocol("WM_DELETE_WINDOW", lambda: self.settings_window.withdraw())
            
            # Main frame
            main_frame = ctk.CTkFrame(self.settings_window)
            main_frame.pack(pady=20, padx=20, fill="both", expand=True)
            
            # Title
            ctk.CTkLabel(main_frame, text="Einstellungen", font=("Arial", 20)).pack(pady=(0, 20))
            
            # Model selection frame
            model_frame = ctk.CTkFrame(main_frame)
            model_frame.pack(fill="x", padx=10, pady=10)
            
            ctk.CTkLabel(model_frame, text="Modell-Auswahl:", font=("Arial", 14)).pack(anchor="w", pady=5)
            
            # Radio buttons for model selection
            self.model_var = ctk.StringVar(value="openai" if self.settings['use_openai'] else "ollama")
            
            ctk.CTkRadioButton(model_frame, text="OpenAI", variable=self.model_var, 
                             value="openai", command=self.update_settings_fields).pack(anchor="w", pady=5)
            ctk.CTkRadioButton(model_frame, text="Ollama", variable=self.model_var, 
                             value="ollama", command=self.update_settings_fields).pack(anchor="w", pady=5)
            
            # OpenAI settings frame
            self.openai_frame = ctk.CTkFrame(main_frame)
            ctk.CTkLabel(self.openai_frame, text="OpenAI Einstellungen:", font=("Arial", 14)).pack(anchor="w", pady=5)
            self.openai_key_entry = ctk.CTkEntry(self.openai_frame, placeholder_text="OpenAI API Key", width=300)
            self.openai_key_entry.pack(pady=10)
            self.openai_key_entry.insert(0, self.settings['openai_key'])
            
            # Ollama settings frame
            self.ollama_frame = ctk.CTkFrame(main_frame)
            ctk.CTkLabel(self.ollama_frame, text="Ollama Einstellungen:", font=("Arial", 14)).pack(anchor="w", pady=5)
            
            self.ollama_url_entry = ctk.CTkEntry(self.ollama_frame, placeholder_text="Ollama URL", width=300)
            self.ollama_url_entry.pack(pady=5)
            self.ollama_url_entry.insert(0, self.settings['ollama_url'])
            
            self.ollama_model_entry = ctk.CTkEntry(self.ollama_frame, placeholder_text="Modell Name", width=300)
            self.ollama_model_entry.pack(pady=5)
            self.ollama_model_entry.insert(0, self.settings['ollama_model'])
            
            # Save button
            ctk.CTkButton(main_frame, text="Speichern", command=self.save_settings_gui).pack(pady=20)
            
            # Initial update of visible fields
            self.update_settings_fields()
        
        # Show window
        self.settings_window.deiconify()
        self.settings_window.lift()
        self.settings_window.focus_force()

    def update_settings_fields(self):
        if self.model_var.get() == "openai":
            self.ollama_frame.pack_forget()
            self.openai_frame.pack(fill="x", padx=10, pady=10)
        else:
            self.openai_frame.pack_forget()
            self.ollama_frame.pack(fill="x", padx=10, pady=10)

    def save_settings_gui(self):
        self.settings['use_openai'] = self.model_var.get() == "openai"
        self.settings['openai_key'] = self.openai_key_entry.get().strip()
        self.settings['ollama_url'] = self.ollama_url_entry.get().strip()
        self.settings['ollama_model'] = self.ollama_model_entry.get().strip()
        
        # Validate settings before saving
        if self.settings['use_openai'] and not self.settings['openai_key']:
            messagebox.showerror("Fehler", "Bitte geben Sie einen OpenAI API Key ein.")
            return
        elif not self.settings['use_openai']:
            if not self.settings['ollama_url']:
                messagebox.showerror("Fehler", "Bitte geben Sie eine Ollama URL ein.")
                return
            if not self.settings['ollama_model']:
                messagebox.showerror("Fehler", "Bitte geben Sie einen Modellnamen ein.")
                return
        
        self.save_settings()
        messagebox.showinfo("Erfolg", "Einstellungen wurden gespeichert!")
        self.settings_window.withdraw()

    def validate_settings(self):
        if self.settings['use_openai']:
            if not self.settings['openai_key']:
                messagebox.showerror("Fehler", "Bitte konfigurieren Sie zuerst Ihren OpenAI API Key in den Einstellungen.")
                self.show_settings()
                return False
        else:
            if not self.settings['ollama_url'] or not self.settings['ollama_model']:
                messagebox.showerror("Fehler", "Bitte konfigurieren Sie zuerst die Ollama-Einstellungen.")
                self.show_settings()
                return False
        return True

    def process_text(self, instruction):
        if not self.validate_settings():
            return
            
        if not self.selected_text:
            messagebox.showerror("Fehler", "Kein Text ausgewählt. Bitte markieren Sie einen Text und drücken Sie Shift-Ctrl-0.")
            return
        
        if self.settings['use_openai']:
            response = self.process_with_openai(instruction, self.selected_text)
        else:
            response = self.process_with_ollama(instruction, self.selected_text)
            
        if response:
            # Copy response to clipboard
            pyperclip.copy(response)
            
            # Minimize our window
            self.root.withdraw()
            
            # Simulate Ctrl+V to paste
            keyboard.send('ctrl+v')
            
            # Wait a moment and restore original clipboard
            time.sleep(0.1)
            pyperclip.copy(self.original_clipboard)

    def process_with_openai(self, instruction, text):
        try:
            # Configure OpenAI client
            client = openai.OpenAI(api_key=self.settings['openai_key'])
            
            # Create chat completion
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": instruction},
                    {"role": "user", "content": text}
                ],
                timeout=30
            )
            
            return response.choices[0].message.content
            
        except openai.AuthenticationError:
            messagebox.showerror("Fehler", "Ungültiger OpenAI API Key. Bitte überprüfen Sie Ihre Einstellungen.")
            self.show_settings()
            return None
        except openai.APITimeoutError:
            messagebox.showerror("Fehler", "Zeitüberschreitung bei der API-Anfrage. Bitte versuchen Sie es erneut.")
            return None
        except openai.APIError as e:
            messagebox.showerror("Fehler", f"OpenAI API Fehler: {str(e)}")
            return None
        except Exception as e:
            messagebox.showerror("Fehler", f"Unerwarteter Fehler: {str(e)}")
            return None

    def process_with_ollama(self, instruction, text):
        prompt = f"{instruction}\n\nText: {text}"
        
        data = {
            'model': self.settings['ollama_model'],
            'prompt': prompt
        }
        
        try:
            response = requests.post(f"{self.settings['ollama_url']}/api/generate",
                                  json=data,
                                  timeout=30)
            
            if response.status_code == 404:
                messagebox.showerror("Fehler", f"Modell '{self.settings['ollama_model']}' nicht gefunden.")
                self.show_settings()
                return None
                
            response.raise_for_status()
            return response.json()['response']
        except requests.exceptions.Timeout:
            messagebox.showerror("Fehler", "Zeitüberschreitung bei der API-Anfrage. Bitte versuchen Sie es erneut.")
            return None
        except requests.exceptions.ConnectionError:
            messagebox.showerror("Fehler", "Verbindung zu Ollama fehlgeschlagen. Ist der Ollama-Server aktiv?")
            return None
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Fehler", f"Ollama API Fehler: {str(e)}")
            return None

    def rephrase(self):
        self.process_text("Formuliere den folgenden Text um, behalte dabei den Inhalt bei, aber ändere die Wortwahl und Struktur:")

    def translate(self):
        self.process_text("Übersetze den folgenden Text ins Deutsche. Wenn der Text bereits auf Deutsch ist, übersetze ihn ins Englische:")

    def summarize(self):
        self.process_text("Fasse den folgenden Text kurz und prägnant zusammen:")

    def rephrase_and_translate(self):
        self.process_text("Formuliere den folgenden Text um und übersetze das Ergebnis. Wenn der Text auf Deutsch ist, übersetze ins Englische, sonst ins Deutsche:")

    def quit_app(self, _=None):
        try:
            if hasattr(self, 'icon'):
                self.icon.stop()
            if hasattr(self, 'root'):
                self.root.destroy()
            os._exit(0)  # Force exit to prevent thread issues
        except:
            os._exit(0)
            
    def show_about(self, _=None):
        about_window = ctk.CTkToplevel(self.root)
        about_window.title("About")
        about_window.geometry("200x150")
        ctk.CTkLabel(about_window, text="Author: Atilla Tenz").pack(pady=5, padx=20)
        ctk.CTkLabel(about_window, text="E-Mail: hunwar@gmail.com").pack(pady=5, padx=20)
        ctk.CTkLabel(about_window, text="Datum: Oktober 2024").pack(pady=5, padx=20)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = TextProcessor()
    app.run()
