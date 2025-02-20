import sys
import webbrowser
import pyttsx3
import speech_recognition as sr
import os
import subprocess
import pywhatkit
import openai
from dotenv import load_dotenv
import winreg
import platform
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTextEdit,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QLabel,
    QProgressBar,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal,QTimer
from PyQt6.QtGui import QFont, QColor, QPalette

# Load environment variables
load_dotenv()

# Initialize OpenAI API
openai.api_key = os.getenv("OPENAI_API_KEY")
if not openai.api_key:
    raise ValueError("OpenAI API key not found. Please set the OPENAI_API_KEY environment variable.")

# Initialize the speech recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()

# Set up wake word
WAKE_WORD = "jarvis"
LISTENING = False

class SpeechEngine:
    """Wrapper for pyttsx3 to handle speech synthesis."""
    def __init__(self):
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        self.engine.setProperty('volume', 1.0)
        self.is_speaking = False

    def speak(self, text):
        """Convert text to speech."""
        if not self.is_speaking:
            self.is_speaking = True
            self.engine.say(text)
            self.engine.runAndWait()
            self.is_speaking = False

    def stop(self):
        """Stop the speech synthesis."""
        self.engine.stop()
        self.is_speaking = False

speech_engine = SpeechEngine()

def find_program_path(program_name):
    """Find program path based on operating system."""
    system = platform.system().lower()
    
    if system == "windows":
        # Common Windows program locations
        common_locations = [
            os.path.join(os.environ.get('ProgramFiles', ''), program_name),
            os.path.join(os.environ.get('ProgramFiles(x86)', ''), program_name),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), program_name),
            os.path.join(os.environ.get('APPDATA', ''), program_name)
        ]
        
        # Add .exe extension if not present
        if not program_name.endswith('.exe'):
            program_name += '.exe'
            
        # Search in common locations
        for location in common_locations:
            if os.path.exists(location):
                return location
            
        # Try Windows Registry
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\\" + program_name) as key:
                return winreg.QueryValue(key, None)
        except WindowsError:
            pass
            
    elif system == "darwin":  # macOS
        common_locations = [
            f"/Applications/{program_name}.app",
            f"{os.path.expanduser('~/Applications')}/{program_name}.app"
        ]
    else:  # Linux
        common_locations = [
            f"/usr/bin/{program_name}",
            f"/usr/local/bin/{program_name}"
        ]
    
    for location in common_locations:
        if os.path.exists(location):
            return location
            
    return None

def open_desktop_app(app_name):
    """Open desktop applications with improved error handling and debugging."""
    try:
        app_name = app_name.lower().strip()
        print(f"Attempting to open: {app_name}")
        
        # Special handling for common applications
        common_apps = {
            "file explorer": {
                "windows": "explorer.exe",
                "darwin": "Finder",
                "linux": "nautilus"
            },
            "microsoft edge": {
                "windows": r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
                "darwin": "/Applications/Microsoft Edge.app",
                "linux": "/usr/bin/microsoft-edge"
            },
            "calculator": {
                "windows": "calc.exe",
                "darwin": "Calculator",
                "linux": "gnome-calculator"
            },
            "clock": {
                "windows": "ms-clock://",
                "darwin": "Clock",
                "linux": "gnome-clocks"
            },
            "command prompt": {
                "windows": "cmd.exe",
                "darwin": "Terminal",
                "linux": "gnome-terminal"
            },
            "settings": {
                "windows": "ms-settings:",
                "darwin": "System Preferences",
                "linux": "gnome-control-center"
            }
        }
        
        system = platform.system().lower()
        
        # Try common apps first
        if app_name in common_apps:
            if system in common_apps[app_name]:
                path = common_apps[app_name][system]
                if os.path.exists(path):
                    print(f"Found common app path: {path}")
                    if system == "darwin":  # macOS
                        subprocess.run(["open", path])
                    else:
                        os.startfile(path) if system == "windows" else subprocess.run([path])
                    speech_engine.speak(f"Opening {app_name}")
                    return
        
        # Try finding the program path
        program_path = find_program_path(app_name)
        if program_path:
            print(f"Found program path: {program_path}")
            if system == "darwin":  # macOS
                subprocess.run(["open", program_path])
            else:
                os.startfile(program_path) if system == "windows" else subprocess.run([program_path])
            speech_engine.speak(f"Opening {app_name}")
            return
        
        # If all else fails, try using the command directly
        try:
            if system == "darwin":  # macOS
                subprocess.run(["open", "-a", app_name])
            elif system == "windows":
                subprocess.run(app_name)
            else:  # Linux
                subprocess.Popen(app_name)
            speech_engine.speak(f"Opening {app_name}")
        except Exception as e:
            print(f"Error running app directly: {e}")
            speech_engine.speak(f"Sorry, I couldn't find or open {app_name}")
            
    except Exception as e:
        print(f"Error opening {app_name}: {e}")
        speech_engine.speak(f"Sorry, I encountered an error while trying to open {app_name}")

def process_command(command, is_voice=True):
    """Process user commands and return response string."""
    try:
        command = command.lower().strip()
        print(f"Processing command: {command}")  # Debug print

        # Website commands
        websites = {
            "open google": "https://google.com",
            "open facebook": "https://facebook.com",
            "open youtube": "https://youtube.com",
            "open instagram": "https://www.instagram.com",
            "open whatsapp": "https://web.whatsapp.com",
            "open github": "https://github.com"
        }

        # Check website commands
        for key, url in websites.items():
            if key in command:
                webbrowser.open(url)
                response = f"Opening {key.replace('open ', '')}"
                if is_voice:
                    speech_engine.speak(response)
                return response

        # Handle song playback
        if command.startswith("play"):
            song_name = command.replace("play", "").strip()
            song_name = " ".join(song_name.split())
            try:
                pywhatkit.playonyt(song_name)
                response = f"Playing {song_name} on YouTube"
                if is_voice:
                    speech_engine.speak(response)
                return response
            except Exception as e:
                response = f"Sorry, I couldn't play that song: {str(e)}"
                if is_voice:
                    speech_engine.speak(response)
                return response

        if "search for" in command or "search" in command:
            if "youtube" in command:
                search_term = command.replace("search for", "").replace("search", "").replace("youtube", "").strip()
                webbrowser.open(f"https://www.youtube.com/results?search_query={search_term}")
                response = f"Searching for {search_term} on YouTube"
            elif "google" in command:
                search_term = command.replace("search for", "").replace("search", "").replace("google", "").strip()
                webbrowser.open(f"https://www.google.com/search?q={search_term}")
                response = f"Searching for {search_term} on Google"
            elif "instagram" in command:
                search_term = command.replace("search for", "").replace("search", "").replace("instagram", "").strip()
                webbrowser.open(f"https://www.instagram.com/explore/people/?q={search_term}")
                response = f"Showing search results for {search_term} on Instagram"
            elif "linkedin" in command:
                search_term = command.replace("search for", "").replace("search", "").replace("linkedin", "").strip()
                webbrowser.open(f"https://www.linkedin.com/search/results/people/?keywords={search_term}")
                response = f"Showing search results for {search_term} on LinkedIn"
            else:
                response = "Please specify a platform to search on (e.g., Google, YouTube, Instagram, LinkedIn)."
            
            if is_voice:
                speech_engine.speak(response)
            return response

        # Open app command
        if command.startswith("open "):
            app_name = command.replace("open ", "").strip()
            response = open_desktop_app(app_name)
            return response

        # Handle ChatGPT responses for everything else
        if command not in ["exit", "quit", "go to sleep"]:
            try:
                response = get_chatbot_response(command)
                if is_voice:
                    speech_engine.speak(response)
                return response
            except Exception as e:
                response = f"Sorry, I couldn't get a response from ChatGPT: {str(e)}"
                if is_voice:
                    speech_engine.speak(response)
                return response

        # Default response if no conditions are met
        response = "I didn't understand that command. Please try again."
        if is_voice:
            speech_engine.speak(response)
        return response

    except Exception as e:
        error_msg = f"Error processing command: {str(e)}"
        if is_voice:
            speech_engine.speak(error_msg)
        return error_msg

def get_chatbot_response(prompt):
    """Get response from OpenAI ChatGPT."""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant named Jarvis."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1000
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"I encountered an error: {str(e)}"

class VoiceThread(QThread):
    """Thread to handle voice commands."""
    command_received = pyqtSignal(str)
    listening_status = pyqtSignal(bool)
    error_occurred = pyqtSignal(str)

    def run(self):
        """Listen for voice commands."""
        global LISTENING
        while True:
            try:
                with sr.Microphone() as source:
                    recognizer.adjust_for_ambient_noise(source, duration=1)
                    self.listening_status.emit(True)
                    print("Listening for wake word...")
                    audio = recognizer.listen(source, timeout=None, phrase_time_limit=5)
                
                try:
                    command = recognizer.recognize_google(audio).lower()
                    print(f"Heard: {command}")
                    
                    if WAKE_WORD in command or LISTENING:
                        if not LISTENING:
                            self.command_received.emit("Wake word detected. How can I help you?")
                            LISTENING = True
                            continue
                        
                        if "stop listening" in command or "go to sleep" in command:
                            LISTENING = False
                            self.command_received.emit("Going to sleep. Say 'Jarvis' to wake me up.")
                            continue
                            
                        self.command_received.emit(command)
                        
                except sr.UnknownValueError:
                    if LISTENING:
                        self.error_occurred.emit("Sorry, I didn't catch that.")
                except sr.RequestError:
                    self.error_occurred.emit("Speech recognition service error.")
                    
            except Exception as e:
                self.error_occurred.emit(f"Error in voice thread: {e}")
                break

class JarvisGUI(QMainWindow):
    """Main GUI window for Jarvis."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Jarvis Assistant")
        self.setGeometry(100, 100, 800, 600)
        self.setup_ui()
        self.setup_voice_thread()

    def setup_ui(self):
        """Set up the user interface."""
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Output text area
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setFont(QFont("Arial", 12))
        self.output_text.setStyleSheet("""
            QTextEdit {
                background-color: #2E3440;
                color: #D8DEE9;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        main_layout.addWidget(self.output_text)

        # Input layout
        input_layout = QHBoxLayout()
        
        # Input field
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Type your command here or click voice command...")
        self.input_field.setFont(QFont("Arial", 12))
        self.input_field.setStyleSheet("""
            QLineEdit {
                background-color: #3B4252;
                color: #ECEFF4;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        input_layout.addWidget(self.input_field)

        # Send button
        self.send_button = QPushButton("Send")
        self.send_button.setFont(QFont("Arial", 12))
        self.send_button.setStyleSheet("""
            QPushButton {
                background-color: #5E81AC;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
        """)
        self.send_button.clicked.connect(self.process_text_command)
        input_layout.addWidget(self.send_button)

        main_layout.addLayout(input_layout)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)

        self.voice_button = QPushButton("🎤 Voice Command")
        self.voice_button.setFont(QFont("Arial", 12))
        self.voice_button.setStyleSheet("""
            QPushButton {
                background-color: #5E81AC;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
        """)
        self.voice_button.clicked.connect(self.toggle_voice_command)
        button_layout.addWidget(self.voice_button)

        self.stop_button = QPushButton("⏹️ Stop")
        self.stop_button.setFont(QFont("Arial", 12))
        self.stop_button.setStyleSheet("""
            QPushButton {
                background-color: #BF616A;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #D08770;
            }
        """)
        self.stop_button.clicked.connect(self.stop_session)
        button_layout.addWidget(self.stop_button)

        self.clear_button = QPushButton("🧹 Clear")
        self.clear_button.setFont(QFont("Arial", 12))
        self.clear_button.setStyleSheet("""
            QPushButton {
                background-color: #BF616A;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #D08770;
            }
        """)
        self.clear_button.clicked.connect(self.clear_output)
        button_layout.addWidget(self.clear_button)

        main_layout.addLayout(button_layout)

        # Loading indicator
        self.loading_indicator = QProgressBar()
        self.loading_indicator.setRange(0, 0)
        self.loading_indicator.setVisible(False)
        main_layout.addWidget(self.loading_indicator)

        # Set main layout
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Set window style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #3B4252;
            }
        """)


    def setup_voice_thread(self):
        """Set up and start the voice thread."""
        self.voice_thread = VoiceThread()
        self.voice_thread.command_received.connect(self.process_voice_command)
        self.voice_thread.error_occurred.connect(self.show_error)
        self.voice_thread.start()

    def process_text_command(self):
        """Process command from text input."""
        command = self.input_field.text().strip()
        self.input_field.clear()
        if command:
            self.output_text.append(f"<span style='color:#88C0D0'>You:</span> {command}")
            self.loading_indicator.setVisible(True)
            QTimer.singleShot(100, lambda: self.process_command(command, is_voice=False))

    def toggle_voice_command(self):
        """Toggle voice command listening."""
        global LISTENING
        LISTENING = not LISTENING
        status = "listening" if LISTENING else "sleeping"
        self.output_text.append(f"<span style='color:#81A1C1'>Jarvis:</span> Voice command is now {status}\n")
        self.voice_button.setText("🎤 Stop Listening" if LISTENING else "🎤 Voice Command")

    def stop_session(self):
        """Stop the current session."""
        global LISTENING
        LISTENING = False
        speech_engine.stop()
        self.output_text.append(f"<span style='color:#81A1C1'>Jarvis:</span> Session stopped.\n")
        self.voice_button.setText("🎤 Voice Command")

    def process_voice_command(self, command):
        """Process command from voice input."""
        self.output_text.append(f"<span style='color:#88C0D0'>You (Voice):</span> {command}")
        self.loading_indicator.setVisible(True)
        QTimer.singleShot(100, lambda: self.process_command(command, is_voice=True))

    def process_command(self, command, is_voice=True):
        """Process the command and update the GUI."""
        try:
            response = process_command(command, is_voice)
            self.output_text.append(f"<span style='color:#81A1C1'>Jarvis:</span> {response}\n")
            self.output_text.verticalScrollBar().setValue(self.output_text.verticalScrollBar().maximum())
            if is_voice:
                speech_engine.speak(response)
        except Exception as e:
            error_msg = f"Error processing command: {str(e)}"
            self.output_text.append(f"<span style='color:#BF616A'>Error:</span> {error_msg}\n")
        finally:
            self.loading_indicator.setVisible(False)

    def show_error(self, message):
        """Display error messages."""
        self.output_text.append(f"<span style='color:#BF616A'>Error:</span> {message}\n")
        self.loading_indicator.setVisible(False)

    def clear_output(self):
        """Clear the output text area."""
        self.output_text.clear()

def main():
    """Main function to run the GUI."""
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(59, 66, 82))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    app.setPalette(palette)
    
    window = JarvisGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
