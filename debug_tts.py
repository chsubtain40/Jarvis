import pyttsx3

print("Initializing TTS Debugger...")
try:
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    print(f"Found {len(voices)} voices.")
    
    for i, voice in enumerate(voices):
        print(f"Testing Voice {i}: {voice.name}")
        try:
            engine.setProperty('voice', voice.id)
            engine.say(f"This is voice number {i}, {voice.name}")
            engine.runAndWait()
            print("  [OK] Speech successful")
        except Exception as e:
            print(f"  [FAIL] Error with this voice: {e}")

except Exception as e:
    print(f"CRITICAL ERROR initializing SAPI5: {e}")

print("Debug complete.")
