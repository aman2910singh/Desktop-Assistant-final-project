import asyncio
import json
import websockets
import speech_recognition as sr
import pyttsx3
import threading
import time
import requests
import webbrowser
import os
import subprocess
import platform
from datetime import datetime
import wikipedia

# --- Corrected handle_voice_input to handle context manager issue ---
def handle_voice_input_thread(assistant, websocket, loop):
    """Handle voice input in a separate thread, communicating with the main event loop."""
    while assistant.is_listening:
        try:
            # Check if the assistant has been commanded to stop listening
            if not assistant.is_listening:
                break
                
            command = assistant.listen()
            
            if command not in ['timeout', 'could not understand'] and not command.startswith('error'):
                asyncio.run_coroutine_threadsafe(
                    websocket.send(json.dumps({
                        'type': 'voice_recognized',
                        'text': command
                    })),
                    loop
                ).result()
                
                response = assistant.process_command(command)
                asyncio.run_coroutine_threadsafe(
                    websocket.send(json.dumps({
                        'type': 'response',
                        'text': response
                    })),
                    loop
                ).result()
                
                assistant.speak(response)
                
                if any(word in command for word in ['goodbye', 'bye', 'exit', 'quit']):
                    assistant.is_listening = False
                    break
            
            # Add a small delay to allow the microphone context to close properly
            time.sleep(0.1)
            
        except Exception as e:
            print(f"Voice input error: {e}")
            assistant.is_listening = False
            # Add a small delay after an error to prevent a tight error loop
            time.sleep(1)

class DesktopAssistant:
    def __init__(self):
        self.tts_engine = pyttsx3.init()
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.is_listening = False
        self.setup_tts()
        
    def setup_tts(self):
        voices = self.tts_engine.getProperty('voices')
        if voices:
            self.tts_engine.setProperty('voice', voices[0].id)
        self.tts_engine.setProperty('rate', 180)
        self.tts_engine.setProperty('volume', 0.9)
    
    def speak(self, text):
        try:
            self.tts_engine.say(text)
            self.tts_engine.runAndWait()
        except Exception as e:
            print(f"TTS Error: {e}")
    
    def listen(self):
        try:
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                # Ensure the microphone is in a "listening" state within the context manager
                print("Listening...")
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
            
            text = self.recognizer.recognize_google(audio)
            return text.lower()
        
        except sr.WaitTimeoutError:
            print("Listening timed out, retrying...")
            return "timeout"
        except sr.UnknownValueError:
            print("Could not understand audio, retrying...")
            return "could not understand"
        except sr.RequestError as e:
            print(f"Speech recognition service error: {e}")
            return f"error: {e}"
        except Exception as e:
            print(f"An unexpected microphone error occurred: {e}")
            return "mic_error"
    
    def get_weather(self, city="London"):
        try:
            api_key = "YOUR_OPENWEATHERMAP_API_KEY"
            if api_key == "YOUR_OPENWEATHERMAP_API_KEY":
                return "Please get a valid OpenWeatherMap API key and replace the placeholder."
            url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
            response = requests.get(url)
            data = response.json()
            if response.status_code == 200:
                temp = data['main']['temp']
                desc = data['weather'][0]['description']
                return f"The weather in {city} is {temp}°C with {desc}."
            else:
                return "Sorry, I couldn't get the weather information for that city."
        except Exception as e:
            print(f"Weather API Error: {e}")
            return "Weather service is currently unavailable."
    
    def search_wikipedia(self, query):
        try:
            summary = wikipedia.summary(query, sentences=2)
            return summary
        except wikipedia.exceptions.DisambiguationError as e:
            return f"Multiple results found. Could you be more specific? Options: {', '.join(e.options[:3])}"
        except wikipedia.exceptions.PageError:
            return "Sorry, I couldn't find information about that topic."
        except Exception as e:
            print(f"Wikipedia Error: {e}")
            return "Wikipedia search is currently unavailable."
    
    def get_time(self):
        now = datetime.now()
        return now.strftime("The current time is %I:%M %p.")
    
    def get_date(self):
        now = datetime.now()
        return now.strftime("Today is %A, %B %d, %Y.")
    
    def open_website(self, website):
        websites = {
            'google': 'https://google.com',
            'youtube': 'https://youtube.com',
            'github': 'https://github.com',
            'stackoverflow': 'https://stackoverflow.com',
            'wikipedia': 'https://wikipedia.org'
        }
        url = websites.get(website.lower())
        if url:
            webbrowser.open(url)
            return f"Opening {website}."
        else:
            return f"I don't know the website {website}."
    
    def open_application(self, app_name):
        app_name = app_name.lower().strip()
        system = platform.system()
        if system == 'Windows':
            apps = {
                'notepad': 'notepad.exe',
                'calculator': 'calc.exe',
                'paint': 'mspaint.exe',
                'task manager': 'taskmgr.exe'
            }
            app = apps.get(app_name)
            if app:
                try:
                    subprocess.Popen(app)
                    return f"Opening {app_name}."
                except FileNotFoundError:
                    return f"Couldn't find the application '{app_name}' on your system."
            else:
                return f"I don't know how to open '{app_name}' on Windows."
        elif system == 'Darwin':
            apps = {
                'notes': 'Notes',
                'calculator': 'Calculator',
                'textedit': 'TextEdit'
            }
            app = apps.get(app_name)
            if app:
                try:
                    subprocess.Popen(['open', '-a', app])
                    return f"Opening {app_name}."
                except FileNotFoundError:
                    return f"Couldn't find the application '{app_name}' on your system."
            else:
                return f"I don't know how to open '{app_name}' on macOS."
        else:
            try:
                subprocess.Popen([app_name])
                return f"Opening {app_name}."
            except FileNotFoundError:
                return f"Couldn't find the application '{app_name}' on your system."
        return f"I don't know how to open {app_name}."
    
    def process_command(self, command):
        command = command.lower().strip()
        if any(greeting in command for greeting in ['hello', 'hi', 'hey']):
            return "Hello! How can I assist you today?"
        elif 'time' in command:
            return self.get_time()
        elif 'date' in command:
            return self.get_date()
        elif 'weather' in command:
            if 'in' in command:
                city = command.split('in')[-1].strip()
                return self.get_weather(city)
            return self.get_weather()
        elif command.startswith('search') or command.startswith('what is'):
            query = command.replace('search', '').replace('what is', '').strip()
            return self.search_wikipedia(query)
        elif command.startswith('open') and any(site in command for site in ['google', 'youtube', 'github', 'stackoverflow', 'wikipedia']):
            for site in ['google', 'youtube', 'github', 'stackoverflow', 'wikipedia']:
                if site in command:
                    return self.open_website(site)
        elif command.startswith('open'):
            app_name = command.replace('open', '').strip()
            return self.open_application(app_name)
        elif 'joke' in command:
            jokes = ["Why don't scientists trust atoms? Because they make up everything!", "I told my wife she was drawing her eyebrows too high. She looked surprised.", "Why don't programmers like nature? It has too many bugs!"]
            import random
            return random.choice(jokes)
        elif any(farewell in command for farewell in ['goodbye', 'bye', 'exit', 'quit']):
            return "Goodbye! Have a great day!"
        else:
            return "I'm sorry, I didn't understand that command. Try asking about time, weather, or say 'open' followed by an application name."

# WebSocket server for communication with frontend
class WebSocketServer:
    def __init__(self, assistant, loop):
        self.assistant = assistant
        self.clients = set()
        self.loop = loop
    
    async def register(self, websocket):
        self.clients.add(websocket)
        print(f"Client connected: {websocket.remote_address}")
    
    async def unregister(self, websocket):
        self.clients.discard(websocket)
        print(f"Client disconnected: {websocket.remote_address}")
    
    async def handle_client(self, websocket, path):
        await self.register(websocket)
        try:
            async for message in websocket:
                data = json.loads(message)
                
                if data['type'] == 'text_command':
                    response = self.assistant.process_command(data['command'])
                    await websocket.send(json.dumps({
                        'type': 'response',
                        'text': response
                    }))
                    threading.Thread(
                        target=self.assistant.speak,
                        args=(response,),
                        daemon=True
                    ).start()
                
                elif data['type'] == 'start_listening':
                    if not self.assistant.is_listening:
                        self.assistant.is_listening = True
                        threading.Thread(
                            target=handle_voice_input_thread,
                            args=(self.assistant, websocket, self.loop,),
                            daemon=True
                        ).start()
                
                elif data['type'] == 'stop_listening':
                    self.assistant.is_listening = False
        
        except websockets.exceptions.ConnectionClosed:
            pass
        finally:
            self.assistant.is_listening = False
            await self.unregister(websocket)

# Main async entry point
async def main():
    print("Starting Desktop Assistant Server...")
    assistant = DesktopAssistant()
    ws_server = WebSocketServer(assistant, asyncio.get_event_loop())
    
    async with websockets.serve(ws_server.handle_client, "localhost", 8765):
        print("Server started at ws://localhost:8765")
        await asyncio.Future()

if __name__ == "__main__":
    asyncio.run(main())