"""Simple Windows voice assistant (Jarvis-like).
Features:
- Text-to-speech (pyttsx3/sapi5)
- Speech recognition via microphone (SpeechRecognition/ Google API)
- Wikipedia summaries, web searches, app launching, time announcements
"""

import pyttsx3  
import speech_recognition as sr  
import datetime
import wikipedia  
import webbrowser
import os
from pathlib import Path
import sys
import time

print("Core process initializing...", flush=True)

# Initialize the text-to-speech engine (Windows SAPI5)
engine = None
try:
    print("Initializing TTS engine...")
    engine = pyttsx3.init("sapi5")
    
    # Get available voices
    voices = engine.getProperty("voices")
    print(f"Found {len(voices)} available voices")
    
    if len(voices) > 0:
        # Try to find a good voice (prefer female voice if available)
        selected_voice = voices[0]  # Default to first voice
        for i, voice in enumerate(voices):
            print(f"Voice {i}: {voice.name} (ID: {voice.id})")
            # Look for a female voice or a more natural sounding voice
            if "female" in voice.name.lower() or "zira" in voice.name.lower() or "hazel" in voice.name.lower():
                selected_voice = voice
                break
        
        engine.setProperty("voice", selected_voice.id)
        print(f"TTS initialized with voice: {selected_voice.name}")
    else:
        print("Warning: No voices found. TTS may not work.")
    
    # Set speech rate and volume for better experience
    engine.setProperty("rate", 150)  # Speed of speech
    engine.setProperty("volume", 1.0)  # Volume level (0.0 to 1.0)
    
    # Test the TTS engine
    print("Testing TTS engine...")
    engine.say("TTS engine is working properly")
    engine.runAndWait()
    print("TTS engine ready!")
    
except Exception as e:
    print(f"Error initializing TTS engine: {e}")
    print("Please ensure you have Windows SAPI5 installed")
    print("You may need to install additional speech voices from Windows Settings")
    engine = None

def speak(text: str) -> None:
    """Speak text aloud using the initialized TTS engine.

    Falls back to printing an error if TTS fails.
    """
    global engine
    print(f"Jarvis: {text}", flush=True)
    
    if engine is None:
        print("[X] TTS engine not initialized. Cannot speak.")
        print("Please check your Windows speech settings and try again.")
        return
    
    try:
        # Add the text to speak (DO NOT use engine.stop() - it breaks pyttsx3 on Windows!)
        engine.say(text)
        
        # Run and wait for speech to complete
        engine.runAndWait()
        
        print("[OK] Speech completed", flush=True)
        
    except RuntimeError as e:
        print(f"[X] Runtime error in speak(): {e}", flush=True)
        print("Text that failed to speak:", text)
        print("Attempting to reinitialize TTS engine...")
        
        # Try to reinitialize the engine
        try:
            # Properly dispose of old engine
            try:
                del engine
            except:
                pass
                
            engine = pyttsx3.init("sapi5", debug=False)
            voices = engine.getProperty("voices")
            if len(voices) > 0:
                # Try to find a good voice
                for voice in voices:
                    if "female" in voice.name.lower() or "zira" in voice.name.lower() or "hazel" in voice.name.lower():
                        engine.setProperty("voice", voice.id)
                        break
                else:
                    engine.setProperty("voice", voices[0].id)
            
            engine.setProperty("rate", 150)
            engine.setProperty("volume", 1.0)
            print("TTS engine reinitialized successfully")
            
            # Try speaking again with the new engine
            engine.say(text)
            engine.runAndWait()
            
        except Exception as reinit_error:
            print(f"Failed to reinitialize TTS: {reinit_error}")
            engine = None
            
    except Exception as e:
        print(f"[X] Unexpected error in speak(): {e}", flush=True)
        print("Text that failed to speak:", text)

def test_tts() -> bool:
    """Test if TTS is working properly.
    
    Returns True if TTS is working, False otherwise.
    """
    if engine is None:
        print("[X] TTS engine is not initialized")
        return False
    
    try:
        print("Testing TTS functionality...")
        engine.say("Testing voice output")
        engine.runAndWait()
        print("[OK] TTS test successful")
        return True
    except Exception as e:
        print(f"[X] TTS test failed: {e}")
        return False

def startup() -> None:
    """Play a short boot-up/initialization sequence and greet the user.

    Includes time-of-day greeting and announces the current time.
    """
    print("=" * 50)
    print("JARVIS VOICE ASSISTANT STARTING UP")
    print("=" * 50)
    
    # Test TTS first
    if not test_tts():
        print("⚠️  TTS is not working properly. Please check your system settings.")
        print("You can still use the assistant, but responses will only be displayed as text.")
        return
    
    speak("Hello I am Jarvis")
    time.sleep(0.5)
    speak("Initializing Jarvis")
    time.sleep(0.5)
    speak("Starting all systems applications")
    time.sleep(0.5)
    speak("Installing and checking all drivers")
    time.sleep(0.5)
    speak("Calibrating and examining all the core processors")
    time.sleep(0.5)
    speak("Checking the internet connection")
    time.sleep(0.5)
    speak("Wait a moment sir")
    time.sleep(0.5)
    speak("All drivers are up and running")
    time.sleep(0.5)
    speak("All systems have been activated")
    time.sleep(0.5)
    speak("Now I am online")
    time.sleep(0.5)

    hour = datetime.datetime.now().hour
    if 0 <= hour <= 12:
        speak("Good morning")
    elif 12 < hour < 18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    time.sleep(0.5)

    c_time = datetime.datetime.now().strftime("%H:%M:%S")
    speak(f"Currently it is {c_time}")
    
    print("=" * 50)
    print("JARVIS IS READY FOR COMMANDS")
    print("=" * 50)

def take_command() -> str | None:
    """Listen from microphone and convert speech to lowercased text.
    Fallback to text input if microphone is unavailable.
    """
    r = sr.Recognizer()
    try:
        # Check if PyAudio is installed/working by attempting to list devices or just try Microphone
        with sr.Microphone() as source:
            print("Listening...")
            # Tune recognition behavior: pauses, energy floor, and auto-adjustment
            r.pause_threshold = 1
            r.energy_threshold = 300
            r.dynamic_energy_threshold = True
            audio = r.listen(source)

        try:
            print("Recognizing...")
            # Use Google's speech recognition with English (India) locale
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}\n")
            return query.lower()
        except Exception:
            print("Say that again please...")
            return None

    except (OSError, AttributeError, ImportError) as e:
        # Fallback for missing PyAudio or no microphone
        print("\n[!] Microphone unavailable (PyAudio not installed or no device found).")
        print(f"Debug Info: {e}")
        print("Switching to text input mode.")
        return input("Type your command: ").lower()

def open_application(app_name: str) -> None:
    """Open common Windows applications by name.

    Tries known installation paths and announces success/failure.
    """
    app_paths = {
        "chrome": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ],
        "notepad": r"C:\Windows\System32\notepad.exe",
        "calculator": r"C:\Windows\System32\calc.exe",
        "paint": r"C:\Windows\System32\mspaint.exe",
        "cmd": r"C:\Windows\System32\cmd.exe",
        "vscode": os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
    }

    try:
        key = app_name.strip()
        if key in app_paths:
            target = app_paths[key]
            if isinstance(target, list):
                for path in target:
                    if Path(path).exists():
                        os.startfile(path)
                        speak(f"Opening {key}")
                        return
                speak(f"Could not find {key} in any of the expected locations")
            else:
                if Path(target).exists():
                    os.startfile(target)
                    speak(f"Opening {key}")
                else:
                    speak(f"Could not find {key} at the expected location")
        else:
            speak(f"Sorry, I don't have the path for {app_name}")
    except Exception as e:
        # Generic fallback to avoid crashing on unexpected errors
        speak(f"Error opening {app_name}")
        print(e)

def search_chrome(query: str) -> None:
    """Open a Google search for the provided voice query.

    Attempts Chrome first; falls back to the default browser.
    """
    try:
        search_query = query.replace("search", "").strip()
        search_url = f"https://www.google.com/search?q={search_query}"
        try:
            webbrowser.get("chrome").open(search_url)
        except webbrowser.Error:
            webbrowser.open(search_url)
        speak(f"Searching for {search_query}")
    except Exception as e:
        print(f"Error during search: {e}")
        speak("Sorry, I encountered an error while searching")

if __name__ == "__main__":
    # Perform the boot sequence and greeting
    startup()

    while True:
        # Block until we capture a speech command (or None on failure)
        query = take_command()

        if query:
            if "wikipedia" in query:
                speak("Searching Wikipedia...")
                try:
                    topic = query.replace("wikipedia", "").strip()
                    results = wikipedia.summary(topic, sentences=2)
                    speak("According to Wikipedia")
                    print(results)
                    speak(results)
                except wikipedia.exceptions.DisambiguationError as e:
                    # Multiple possible pages; ask the user to refine
                    speak("There are multiple results for this query. Please be more specific.")
                    print("Possible matches:", e.options[:5])
                except wikipedia.exceptions.PageError:
                    speak("Sorry, I couldn't find any information about that on Wikipedia.")
                except Exception as e:
                    speak("Sorry, there was an error while searching Wikipedia.")
                    print(e)

            elif "search" in query:
                search_chrome(query)

            elif "open" in query:
                app_name = query.replace("open", "").strip()
                if "youtube" in app_name:
                    try:
                        webbrowser.get("chrome").open("https://www.youtube.com")
                    except webbrowser.Error:
                        webbrowser.open("https://www.youtube.com")
                    speak("Opening YouTube")
                elif "google" in app_name:
                    try:
                        webbrowser.get("chrome").open("https://www.google.com")
                    except webbrowser.Error:
                        webbrowser.open("https://www.google.com")
                    speak("Opening Google")
                else:
                    open_application(app_name)

            elif "the time" in query:
                str_time = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"The time is {str_time}")
            
            elif "test voice" in query or "test speech" in query:
                speak("Testing voice output. Can you hear me clearly?")
                print("Voice test completed. If you can't hear the voice, check your system audio settings.")

            elif "voice status" in query or "tts status" in query:
                if engine is not None:
                    voices = engine.getProperty("voices")
                    current_voice = engine.getProperty("voice")
                    rate = engine.getProperty("rate")
                    volume = engine.getProperty("volume")
                    speak("Voice system is active and working properly")
                    print(f"Current voice: {current_voice}")
                    print(f"Speech rate: {rate}")
                    print(f"Volume: {volume}")
                    print(f"Available voices: {len(voices)}")
                else:
                    speak("Voice system is not working")
                    print("[X] TTS engine is not initialized")

            elif any(word in query for word in ["exit", "quit", "goodbye"]):
                print("Debug: Exit condition triggered")
                speak("Goodbye!")
                sys.exit(0)