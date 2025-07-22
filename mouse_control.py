import speech_recognition as sr
import pyttsx3
import pyautogui
import pywhatkit
import webbrowser
import os
import time
import datetime
import wikipedia
import requests
import json
from fuzzywuzzy import fuzz
from plyer import notification
import psutil
import numpy as np
import screen_brightness_control as sbc
import subprocess
from threading import Thread
import keyboard
import wolframalpha
import pyjokes
import speedtest
import cv2
import pyperclip
import smtplib
from email.message import EmailMessage
import clipboard
import sys
import screeninfo
import random
import calendar
import socket
import platform
from bs4 import BeautifulSoup
import winshell
import win32com.client
import win32con
import win32gui
import urllib.parse
from urllib.request import urlopen
import instaloader
import tweepy
from geopy.geocoders import Nominatim
from timezonefinder import TimezoneFinder
from datetime import datetime as dt
import pytz
from newsapi import NewsApiClient
import pyowm
from pyowm.utils.config import get_default_config

# ===================== INITIALIZATION =====================
engine = pyttsx3.init()
recognizer = sr.Recognizer()
microphone = sr.Microphone()

# Audio tuning
recognizer.dynamic_energy_threshold = True
recognizer.energy_threshold = 250
recognizer.pause_threshold = 1.2
recognizer.phrase_threshold = 0.3

# Voice settings
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Female voice
engine.setProperty('rate', 160)
engine.setProperty('volume', 1.0)

# Global states
assistant_active = False
last_command = ""
command_history = []
media_playing = False
current_volume = 50
wake_phrases = ["ok jarvis", "hey jarvis", "jarvis", "okay jarvis"]
running = True

# API Keys (replace with your own)
NEWS_API_KEY = 'your_news_api_key'
WEATHER_API_KEY = 'your_weather_api_key'
WOLFRAM_ALPHA_APP_ID = 'your_wolframalpha_app_id'

# Initialize APIs
newsapi = NewsApiClient(api_key=NEWS_API_KEY)
weather_config = get_default_config()
weather_config['language'] = 'en'
owm = pyowm.OWM(WEATHER_API_KEY, weather_config)
wolfram_client = wolframalpha.Client(WOLFRAM_ALPHA_APP_ID)

# ===================== UTILITY FUNCTIONS =====================
def speak(text, priority='normal'):
    """Enhanced speech with priority levels"""
    print(f"🎤 [{priority.upper()}] {text}")
    engine.say(text)
    engine.runAndWait()
    command_history.append(f"Assistant: {text}")

def adjust_for_ambient_noise():
    """Advanced microphone calibration"""
    with microphone as source:
        print("🔊 Calibrating microphone...")
        recognizer.adjust_for_ambient_noise(source, duration=3)
        recognizer.energy_threshold = max(recognizer.energy_threshold * 0.5, 100)

def listen(timeout=5, phrase_time_limit=10):
    """Improved listen function with error handling"""
    global last_command
    
    with microphone as source:
        try:
            print(f"🎧 Listening (Sensitivity: {recognizer.energy_threshold:.1f})...")
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
            
            try:
                command = recognizer.recognize_google(audio, language='en-in')
                print(f"🗣️ You said: {command}")
                last_command = command.lower()
                command_history.append(f"User: {command}")
                return last_command
            except sr.UnknownValueError:
                if assistant_active:
                    speak("Could you repeat that please?", 'low')
                return ""
            except sr.RequestError:
                speak("Having trouble connecting to the internet", 'high')
                return ""
                
        except sr.WaitTimeoutError:
            return ""
        except Exception as e:
            print(f"Listening error: {e}")
            return ""

def fuzzy_match(command, options, min_score=65):
    """Improved fuzzy matching"""
    best_match = None
    highest_score = 0
    
    command = command.lower().strip()
    
    for option in options:
        option = option.lower()
        full_score = fuzz.ratio(command, option)
        partial_score = fuzz.partial_ratio(command, option)
        token_score = fuzz.token_sort_ratio(command, option)
        score = max(full_score, partial_score, token_score)
        
        if score > highest_score and score >= min_score:
            highest_score = score
            best_match = option
    
    if highest_score >= min_score:
        print(f"🔍 Matched '{command}' to '{best_match}' (score: {highest_score})")
        return best_match
    return None

def confirm_action(action):
    """Smart confirmation system"""
    responses = [
        f"Did you say '{action}'?",
        f"Confirm: You want {action}?",
        f"Should I proceed with {action}?"
    ]
    speak(np.random.choice(responses), 'normal')
    response = listen()
    return response and any(word in response for word in ["yes", "yeah", "correct", "do it"])

def check_wake_phrase(command):
    """Check if the command contains any wake phrase"""
    return any(wake_phrase in command for wake_phrase in wake_phrases)

def extract_command(command):
    """Extract the actual command after the wake phrase"""
    for wake_phrase in wake_phrases:
        if wake_phrase in command:
            return command.replace(wake_phrase, "").strip()
    return command.strip()

def listen_for_wake_word():
    """Continuously listens for the wake word"""
    global assistant_active
    print("🔇 Sleeping... Waiting for wake word...")
    
    while running and not assistant_active:
        command = listen(timeout=None)  # No timeout - continuous listening
        if command and any(wake_phrase in command for wake_phrase in wake_phrases):
            assistant_active = True
            speak("Yes Giri, how can I help you?", 'normal')
            return extract_command(command)
        time.sleep(0.1)

def type_text(text):
    """Type text with human-like delays"""
    for char in text:
        pyautogui.write(char)
        time.sleep(random.uniform(0.05, 0.15))

# ===================== ENHANCED COMMAND HANDLERS =====================

def handle_youtube_commands(command):
    """Enhanced YouTube control with search and navigation"""
    youtube_actions = {
        "search": ["search for", "look up", "find"],
        "subscribe": ["subscribe", "follow channel"],
        "like": ["like video", "thumbs up"],
        "dislike": ["dislike video", "thumbs down"],
        "comment": ["add comment", "post comment"],
        "history": ["watch history", "view history"],
        "library": ["my library", "youtube library"],
        "trending": ["trending videos", "popular videos"],
        "subscriptions": ["my subscriptions", "subscribed channels"],
        "playlist": ["my playlist", "saved playlist"],
        "settings": ["youtube settings", "account settings"],
        "fullscreen": ["full screen", "maximize"],
        "theater mode": ["theater mode", "cinema mode"],
        "miniplayer": ["mini player", "small window"],
        "speed": ["playback speed", "change speed"],
        "quality": ["video quality", "change quality"],
        "captions": ["subtitles", "closed captions"],
        "autoplay": ["auto play", "next video"],
        "loop": ["repeat video", "loop video"]
    }
    
    # First check for media control commands (play/pause, volume, etc.)
    if handle_media_control_commands(command):
        return True
    
    # Handle specific YouTube actions
    for action, keywords in youtube_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "search":
                    query = command.replace("search for", "").replace("look up", "").replace("find", "").strip()
                    if not query:
                        speak("What would you like to search for on YouTube?", 'normal')
                        query = listen()
                    if query:
                        # Click on search bar
                        pyautogui.click(600, 120)
                        time.sleep(0.5)
                        # Type the query
                        type_text(query)
                        time.sleep(0.5)
                        # Press enter
                        keyboard.press_and_release('enter')
                        speak(f"Searching YouTube for {query}", 'normal')
                elif action == "subscribe":
                    pyautogui.click(950, 550)  # Approximate subscribe button location
                    speak("Subscribed to the channel", 'normal')
                elif action == "like":
                    pyautogui.click(800, 850)  # Approximate like button location
                    speak("Liked the video", 'normal')
                elif action == "dislike":
                    pyautogui.click(850, 850)  # Approximate dislike button location
                    speak("Disliked the video", 'normal')
                elif action == "comment":
                    pyautogui.click(950, 900)  # Approximate comment box location
                    time.sleep(1)
                    speak("What would you like to comment?", 'normal')
                    comment = listen()
                    if comment:
                        type_text(comment)
                        keyboard.press_and_release('tab')
                        keyboard.press_and_release('enter')
                        speak("Comment posted", 'normal')
                elif action == "history":
                    webbrowser.open("https://www.youtube.com/feed/history")
                    speak("Opening watch history", 'normal')
                elif action == "library":
                    webbrowser.open("https://www.youtube.com/library")
                    speak("Opening your YouTube library", 'normal')
                elif action == "trending":
                    webbrowser.open("https://www.youtube.com/feed/trending")
                    speak("Showing trending videos", 'normal')
                elif action == "subscriptions":
                    webbrowser.open("https://www.youtube.com/feed/subscriptions")
                    speak("Showing your subscriptions", 'normal')
                elif action == "playlist":
                    webbrowser.open("https://www.youtube.com/playlist")
                    speak("Opening your playlists", 'normal')
                elif action == "settings":
                    pyautogui.click(1800, 100)  # Profile icon
                    time.sleep(0.5)
                    pyautogui.click(1800, 300)  # Settings option
                    speak("Opening YouTube settings", 'normal')
                elif action == "fullscreen":
                    keyboard.press_and_release('f')
                    speak("Toggling fullscreen mode", 'normal')
                elif action == "theater mode":
                    keyboard.press_and_release('t')
                    speak("Toggling theater mode", 'normal')
                elif action == "miniplayer":
                    keyboard.press_and_release('i')
                    speak("Toggling miniplayer", 'normal')
                elif action == "speed":
                    pyautogui.rightClick(900, 800)  # Right click on video
                    time.sleep(0.5)
                    pyautogui.click(900, 650)  # Speed menu
                    speak("Current playback speed options displayed", 'normal')
                elif action == "quality":
                    pyautogui.rightClick(900, 800)  # Right click on video
                    time.sleep(0.5)
                    pyautogui.click(900, 600)  # Quality menu
                    speak("Current video quality options displayed", 'normal')
                elif action == "captions":
                    keyboard.press_and_release('c')
                    speak("Toggling captions", 'normal')
                elif action == "autoplay":
                    # Location of autoplay toggle may vary
                    pyautogui.click(1400, 900)
                    speak("Toggling autoplay", 'normal')
                elif action == "loop":
                    # Right click on video and select loop
                    pyautogui.rightClick(900, 800)
                    time.sleep(0.5)
                    pyautogui.click(900, 550)
                    speak("Toggling loop", 'normal')
                return True
            except Exception as e:
                print(f"YouTube control error: {e}")
                speak("Couldn't complete that YouTube action", 'high')
                return True
    
    # Handle YouTube play commands
    if "youtube" in command.lower() or "video" in command.lower():
        query = command.replace("on youtube", "").replace("play", "").replace("video", "").strip()
        if not query:
            speak("What would you like to watch on YouTube?", 'normal')
            query = listen()
        if query:
            speak(f"Playing {query} on YouTube", 'normal')
            pywhatkit.playonyt(query)
            return True
    
    return False

def handle_whatsapp_commands(command):
    """Enhanced WhatsApp control with chat and call functionality"""
    whatsapp_actions = {
        "new chat": ["new chat", "start conversation", "message someone"],
        "search chat": ["search chat", "find conversation"],
        "send message": ["send message", "text someone"],
        "make call": ["make call", "voice call", "call someone"],
        "video call": ["video call", "video chat"],
        "status": ["view status", "check status"],
        "archive": ["archive chat", "hide chat"],
        "mute": ["mute chat", "silence notifications"],
        "star": ["star message", "save message"],
        "delete": ["delete chat", "remove conversation"],
        "clear": ["clear chat", "delete messages"],
        "settings": ["whatsapp settings", "account settings"],
        "logout": ["logout", "sign out"],
        "group": ["create group", "new group"],
        "broadcast": ["broadcast message", "send to many"]
    }
    
    if not fuzzy_match(command, ["whatsapp", "message", "chat"]):
        return False
    
    for action, keywords in whatsapp_actions.items():
        if fuzzy_match(command, keywords):
            try:
                webbrowser.open("https://web.whatsapp.com")
                time.sleep(5)  # Wait for WhatsApp to load
                
                if action == "new chat":
                    pyautogui.click(200, 200)  # New chat icon
                    time.sleep(1)
                    speak("Who would you like to message?", 'normal')
                    contact = listen()
                    if contact:
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        speak(f"Opened chat with {contact}", 'normal')
                elif action == "search chat":
                    pyautogui.click(300, 150)  # Search bar
                    time.sleep(0.5)
                    speak("Which chat would you like to find?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "send message":
                    speak("Who would you like to message?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        speak("What should I send?", 'normal')
                        message = listen()
                        if message:
                            type_text(message)
                            keyboard.press_and_release('enter')
                            speak("Message sent", 'normal')
                elif action == "make call":
                    speak("Who would you like to call?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1300, 150)  # Call button
                        speak(f"Calling {contact}", 'normal')
                elif action == "video call":
                    speak("Who would you like to video call?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1350, 150)  # Video call button
                        speak(f"Starting video call with {contact}", 'normal')
                elif action == "status":
                    pyautogui.click(150, 150)  # Status tab
                    speak("Showing status updates", 'normal')
                elif action == "archive":
                    speak("Which chat would you like to archive?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1800, 250)  # More options
                        time.sleep(0.5)
                        pyautogui.click(1800, 350)  # Archive option
                        speak(f"Archived chat with {contact}", 'normal')
                elif action == "mute":
                    speak("Which chat would you like to mute?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1800, 250)  # More options
                        time.sleep(0.5)
                        pyautogui.click(1800, 400)  # Mute option
                        speak(f"Muted chat with {contact}", 'normal')
                elif action == "star":
                    speak("Please open the message you want to star", 'normal')
                    time.sleep(2)
                    pyautogui.rightClick(900, 600)  # Right click on message
                    time.sleep(0.5)
                    pyautogui.click(900, 550)  # Star option
                    speak("Message starred", 'normal')
                elif action == "delete":
                    speak("Which chat would you like to delete?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1800, 250)  # More options
                        time.sleep(0.5)
                        pyautogui.click(1800, 450)  # Delete option
                        time.sleep(0.5)
                        pyautogui.click(1000, 700)  # Confirm delete
                        speak(f"Deleted chat with {contact}", 'normal')
                elif action == "clear":
                    speak("Which chat would you like to clear?", 'normal')
                    contact = listen()
                    if contact:
                        # Search for contact
                        pyautogui.click(300, 150)
                        type_text(contact)
                        time.sleep(1)
                        keyboard.press_and_release('enter')
                        time.sleep(1)
                        pyautogui.click(1800, 250)  # More options
                        time.sleep(0.5)
                        pyautogui.click(1800, 500)  # Clear chat option
                        speak(f"Cleared chat with {contact}", 'normal')
                elif action == "settings":
                    pyautogui.click(1800, 100)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(1800, 200)  # Settings option
                    speak("Opening WhatsApp settings", 'normal')
                elif action == "logout":
                    pyautogui.click(1800, 100)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(1800, 800)  # Logout option
                    speak("Logged out of WhatsApp", 'normal')
                elif action == "group":
                    pyautogui.click(200, 200)  # New chat icon
                    time.sleep(0.5)
                    pyautogui.click(200, 250)  # New group option
                    speak("Creating new group. Who should be in the group?", 'normal')
                elif action == "broadcast":
                    pyautogui.click(200, 200)  # New chat icon
                    time.sleep(0.5)
                    pyautogui.click(200, 300)  # New broadcast option
                    speak("Creating broadcast list. Who should receive the message?", 'normal')
                return True
            except Exception as e:
                print(f"WhatsApp control error: {e}")
                speak("Couldn't complete that WhatsApp action", 'high')
                return True
    
    # Default WhatsApp opening
    if fuzzy_match(command, ["whatsapp", "open whatsapp"]):
        webbrowser.open("https://web.whatsapp.com")
        speak("Opening WhatsApp Web", 'normal')
        return True
    
    return False

def handle_linkedin_commands(command):
    """Enhanced LinkedIn control with profile and job search"""
    linkedin_actions = {
        "profile": ["view profile", "my profile"],
        "network": ["my network", "connections"],
        "jobs": ["find jobs", "job search"],
        "notifications": ["view notifications", "check alerts"],
        "messages": ["view messages", "check chats"],
        "post": ["create post", "share update"],
        "search": ["search linkedin", "find people"],
        "settings": ["linkedin settings", "account settings"],
        "logout": ["logout", "sign out"],
        "feed": ["news feed", "home page"],
        "company": ["search company", "find organization"],
        "learning": ["linkedin learning", "courses"],
        "groups": ["view groups", "my groups"],
        "events": ["find events", "upcoming events"]
    }
    
    if not fuzzy_match(command, ["linkedin", "professional network"]):
        return False
    
    for action, keywords in linkedin_actions.items():
        if fuzzy_match(command, keywords):
            try:
                webbrowser.open("https://www.linkedin.com")
                time.sleep(5)  # Wait for LinkedIn to load
                
                if action == "profile":
                    pyautogui.click(300, 150)  # Profile menu
                    time.sleep(0.5)
                    pyautogui.click(300, 200)  # View profile
                    speak("Opening your profile", 'normal')
                elif action == "network":
                    pyautogui.click(500, 150)  # Network tab
                    speak("Showing your network connections", 'normal')
                elif action == "jobs":
                    pyautogui.click(700, 150)  # Jobs tab
                    speak("Showing job listings", 'normal')
                elif action == "notifications":
                    pyautogui.click(1100, 150)  # Notifications bell
                    speak("Showing your notifications", 'normal')
                elif action == "messages":
                    pyautogui.click(1200, 150)  # Messages icon
                    speak("Opening your messages", 'normal')
                elif action == "post":
                    pyautogui.click(900, 600)  # Start a post
                    speak("What would you like to post?", 'normal')
                    post_text = listen()
                    if post_text:
                        type_text(post_text)
                        speak("Ready to post. Should I share it?", 'normal')
                        confirm = listen()
                        if confirm and "yes" in confirm:
                            pyautogui.click(900, 800)  # Post button
                            speak("Post shared", 'normal')
                        else:
                            pyautogui.click(1000, 250)  # Close post
                            speak("Post discarded", 'normal')
                elif action == "search":
                    pyautogui.click(400, 100)  # Search bar
                    time.sleep(0.5)
                    speak("What would you like to search for?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        keyboard.press_and_release('enter')
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "settings":
                    pyautogui.click(1800, 150)  # Me menu
                    time.sleep(0.5)
                    pyautogui.click(1800, 300)  # Settings & Privacy
                    speak("Opening LinkedIn settings", 'normal')
                elif action == "logout":
                    pyautogui.click(1800, 150)  # Me menu
                    time.sleep(0.5)
                    pyautogui.click(1800, 800)  # Sign out
                    speak("Logged out of LinkedIn", 'normal')
                elif action == "feed":
                    pyautogui.click(200, 150)  # Home tab
                    speak("Showing your news feed", 'normal')
                elif action == "company":
                    pyautogui.click(400, 100)  # Search bar
                    time.sleep(0.5)
                    speak("Which company would you like to find?", 'normal')
                    company = listen()
                    if company:
                        type_text(company)
                        keyboard.press_and_release('enter')
                        speak(f"Searching for {company}", 'normal')
                elif action == "learning":
                    webbrowser.open("https://www.linkedin.com/learning")
                    speak("Opening LinkedIn Learning", 'normal')
                elif action == "groups":
                    webbrowser.open("https://www.linkedin.com/groups")
                    speak("Showing your groups", 'normal')
                elif action == "events":
                    webbrowser.open("https://www.linkedin.com/events")
                    speak("Showing upcoming events", 'normal')
                return True
            except Exception as e:
                print(f"LinkedIn control error: {e}")
                speak("Couldn't complete that LinkedIn action", 'high')
                return True
    
    # Default LinkedIn opening
    if fuzzy_match(command, ["linkedin", "open linkedin"]):
        webbrowser.open("https://www.linkedin.com")
        speak("Opening LinkedIn", 'normal')
        return True
    
    return False

def handle_chatgpt_commands(command):
    """Enhanced ChatGPT interaction"""
    chatgpt_actions = {
        "new chat": ["new chat", "start over", "clear conversation"],
        "search": ["search chat", "find in chat"],
        "copy": ["copy response", "save answer"],
        "share": ["share chat", "export conversation"],
        "settings": ["chatgpt settings", "preferences"],
        "feedback": ["give feedback", "rate response"],
        "examples": ["show examples", "suggest prompts"],
        "upgrade": ["upgrade account", "go pro"],
        "help": ["chatgpt help", "how to use"]
    }
    
    if not fuzzy_match(command, ["chatgpt", "openai", "ai assistant"]):
        return False
    
    for action, keywords in chatgpt_actions.items():
        if fuzzy_match(command, keywords):
            try:
                webbrowser.open("https://chat.openai.com")
                time.sleep(5)  # Wait for ChatGPT to load
                
                if action == "new chat":
                    pyautogui.click(200, 150)  # New chat button
                    speak("Started new conversation", 'normal')
                elif action == "search":
                    pyautogui.click(1800, 100)  # Search chats
                    time.sleep(0.5)
                    speak("What would you like to search for?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "copy":
                    pyautogui.rightClick(900, 600)  # Right click on response
                    time.sleep(0.5)
                    pyautogui.click(900, 550)  # Copy option
                    speak("Response copied to clipboard", 'normal')
                elif action == "share":
                    pyautogui.click(1700, 150)  # Share chat
                    speak("Preparing to share this conversation", 'normal')
                elif action == "settings":
                    pyautogui.click(1800, 100)  # Settings
                    speak("Opening ChatGPT settings", 'normal')
                elif action == "feedback":
                    pyautogui.click(1600, 600)  # Feedback buttons
                    speak("Please select your feedback", 'normal')
                elif action == "examples":
                    pyautogui.click(900, 800)  # Example prompts
                    speak("Here are some example prompts", 'normal')
                elif action == "upgrade":
                    pyautogui.click(1800, 100)  # Upgrade button
                    speak("Showing upgrade options", 'normal')
                elif action == "help":
                    webbrowser.open("https://help.openai.com")
                    speak("Opening ChatGPT help center", 'normal')
                return True
            except Exception as e:
                print(f"ChatGPT control error: {e}")
                speak("Couldn't complete that ChatGPT action", 'high')
                return True
    
    # Handle ChatGPT queries
    if fuzzy_match(command, ["ask chatgpt", "question for ai"]):
        query = command.replace("ask chatgpt", "").replace("question for ai", "").strip()
        if not query:
            speak("What would you like to ask ChatGPT?", 'normal')
            query = listen()
        if query:
            webbrowser.open("https://chat.openai.com")
            time.sleep(5)
            pyautogui.click(900, 800)  # Message box
            type_text(query)
            keyboard.press_and_release('enter')
            speak("Getting response from ChatGPT", 'normal')
            return True
    
    # Default ChatGPT opening
    if fuzzy_match(command, ["chatgpt", "open chatgpt"]):
        webbrowser.open("https://chat.openai.com")
        speak("Opening ChatGPT", 'normal')
        return True
    
    return False

def handle_notepad_commands(command):
    """Enhanced Notepad operations"""
    notepad_actions = {
        "new": ["new file", "clear document"],
        "open": ["open file", "load document"],
        "save": ["save file", "save document"],
        "save as": ["save as", "save copy"],
        "close": ["close file", "exit document"],
        "print": ["print file", "print document"],
        "find": ["find text", "search document"],
        "replace": ["replace text", "find and replace"],
        "font": ["change font", "text style"],
        "zoom": ["zoom in", "zoom out"],
        "word wrap": ["toggle word wrap", "enable word wrap"],
        "status bar": ["toggle status bar", "show status"],
        "time date": ["insert timestamp", "current time"]
    }
    
    if not fuzzy_match(command, ["notepad", "text editor"]):
        return False
    
    # Open Notepad if not already open
    try:
        os.system("start notepad")
        time.sleep(2)  # Wait for Notepad to open
    except:
        speak("Couldn't open Notepad", 'high')
        return True
    
    for action, keywords in notepad_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "new":
                    keyboard.press_and_release('ctrl+n')
                    speak("Created new document", 'normal')
                elif action == "open":
                    keyboard.press_and_release('ctrl+o')
                    speak("Opening file dialog", 'normal')
                elif action == "save":
                    keyboard.press_and_release('ctrl+s')
                    speak("File saved", 'normal')
                elif action == "save as":
                    keyboard.press_and_release('ctrl+shift+s')
                    speak("Save as dialog opened", 'normal')
                elif action == "close":
                    keyboard.press_and_release('ctrl+w')
                    speak("Document closed", 'normal')
                elif action == "print":
                    keyboard.press_and_release('ctrl+p')
                    speak("Print dialog opened", 'normal')
                elif action == "find":
                    keyboard.press_and_release('ctrl+f')
                    speak("Find dialog opened. What should I search for?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "replace":
                    keyboard.press_and_release('ctrl+h')
                    speak("Replace dialog opened", 'normal')
                elif action == "font":
                    keyboard.press_and_release('ctrl+shift+f')
                    speak("Font dialog opened", 'normal')
                elif action == "zoom":
                    if "in" in command:
                        keyboard.press_and_release('ctrl+plus')
                        speak("Zoomed in", 'normal')
                    else:
                        keyboard.press_and_release('ctrl+minus')
                        speak("Zoomed out", 'normal')
                elif action == "word wrap":
                    keyboard.press_and_release('ctrl+shift+w')
                    speak("Toggled word wrap", 'normal')
                elif action == "status bar":
                    # Notepad doesn't have a direct shortcut for this
                    pyautogui.click(50, 50)  # View menu
                    time.sleep(0.5)
                    pyautogui.click(50, 300)  # Status bar option
                    speak("Toggled status bar", 'normal')
                elif action == "time date":
                    keyboard.press_and_release('f5')
                    speak("Inserted current time and date", 'normal')
                return True
            except Exception as e:
                print(f"Notepad control error: {e}")
                speak("Couldn't complete that Notepad action", 'high')
                return True
    
    # Handle text input
    if fuzzy_match(command, ["type", "write"]):
        text = command.replace("type", "").replace("write", "").strip()
        if not text:
            speak("What would you like me to type?", 'normal')
            text = listen()
        if text:
            type_text(text)
            speak("Text entered", 'normal')
            return True
    
    # Default Notepad opening
    if fuzzy_match(command, ["notepad", "open notepad"]):
        speak("Notepad is already open", 'normal')
        return True
    
    return False

def handle_calculator_commands(command):
    """Enhanced Calculator operations"""
    calc_actions = {
        "standard": ["standard mode", "basic calculator"],
        "scientific": ["scientific mode", "advanced calculator"],
        "programmer": ["programmer mode", "hex calculator"],
        "date calc": ["date calculation", "date difference"],
        "currency": ["currency converter", "money exchange"],
        "volume": ["volume converter", "liquid measurement"],
        "length": ["length converter", "distance measurement"],
        "weight": ["weight converter", "mass measurement"],
        "temperature": ["temperature converter", "degrees conversion"],
        "energy": ["energy converter", "power measurement"],
        "area": ["area converter", "size measurement"],
        "speed": ["speed converter", "velocity measurement"],
        "time": ["time converter", "clock conversion"],
        "memory": ["memory store", "save number"],
        "recall": ["memory recall", "get saved number"],
        "clear": ["clear all", "reset calculator"]
    }
    
    if not fuzzy_match(command, ["calculator", "calc"]):
        return False
    
    # Open Calculator if not already open
    try:
        os.system("start calc")
        time.sleep(2)  # Wait for Calculator to open
    except:
        speak("Couldn't open Calculator", 'high')
        return True
    
    for action, keywords in calc_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "standard":
                    keyboard.press_and_release('ctrl+1')
                    speak("Standard calculator mode", 'normal')
                elif action == "scientific":
                    keyboard.press_and_release('ctrl+2')
                    speak("Scientific calculator mode", 'normal')
                elif action == "programmer":
                    keyboard.press_and_release('ctrl+3')
                    speak("Programmer calculator mode", 'normal')
                elif action == "date calc":
                    keyboard.press_and_release('ctrl+4')
                    speak("Date calculation mode", 'normal')
                elif action == "currency":
                    # Calculator doesn't have direct shortcuts for converters
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 300)  # Currency converter
                    speak("Currency converter opened", 'normal')
                elif action == "volume":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 350)  # Volume converter
                    speak("Volume converter opened", 'normal')
                elif action == "length":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 400)  # Length converter
                    speak("Length converter opened", 'normal')
                elif action == "weight":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 450)  # Weight converter
                    speak("Weight converter opened", 'normal')
                elif action == "temperature":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 500)  # Temperature converter
                    speak("Temperature converter opened", 'normal')
                elif action == "energy":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 550)  # Energy converter
                    speak("Energy converter opened", 'normal')
                elif action == "area":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 600)  # Area converter
                    speak("Area converter opened", 'normal')
                elif action == "speed":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 650)  # Speed converter
                    speak("Speed converter opened", 'normal')
                elif action == "time":
                    pyautogui.click(50, 50)  # Menu button
                    time.sleep(0.5)
                    pyautogui.click(50, 700)  # Time converter
                    speak("Time converter opened", 'normal')
                elif action == "memory":
                    # Store current value in memory
                    keyboard.press_and_release('ctrl+m')
                    speak("Value stored in memory", 'normal')
                elif action == "recall":
                    # Recall memory value
                    keyboard.press_and_release('ctrl+r')
                    speak("Recalled value from memory", 'normal')
                elif action == "clear":
                    keyboard.press_and_release('esc')
                    speak("Calculator cleared", 'normal')
                return True
            except Exception as e:
                print(f"Calculator control error: {e}")
                speak("Couldn't complete that Calculator action", 'high')
                return True
    
    # Handle calculations
    if any(op in command for op in ["plus", "minus", "times", "divided by", "square root", "power of"]):
        try:
            # Map spoken terms to calculator buttons
            calc_map = {
                "plus": "+",
                "minus": "-",
                "times": "*",
                "multiplied by": "*",
                "divided by": "/",
                "square root": "sqrt(",
                "power of": "^",
                "percent": "%"
            }
            
            # Replace words with symbols
            for word, symbol in calc_map.items():
                command = command.replace(word, symbol)
            
            # Handle special cases
            if "sqrt" in command:
                command = command.replace("sqrt", "sqrt(") + ")"
            
            # Type the calculation
            pyautogui.click(500, 500)  # Focus on calculator
            type_text(command)
            keyboard.press_and_release('enter')
            
            # Get result from clipboard
            result = pyperclip.paste()
            speak(f"The result is {result}", 'normal')
            return True
        except Exception as e:
            print(f"Calculation error: {e}")
            speak("Couldn't complete that calculation", 'high')
            return True
    
    # Default Calculator opening
    if fuzzy_match(command, ["calculator", "open calculator"]):
        speak("Calculator is already open", 'normal')
        return True
    
    return False

def handle_file_explorer_commands(command):
    """Enhanced File Explorer operations"""
    explorer_actions = {
        "new folder": ["create folder", "make directory"],
        "rename": ["rename file", "change name"],
        "delete": ["delete file", "remove item"],
        "properties": ["file properties", "show details"],
        "open": ["open file", "launch item"],
        "copy": ["copy file", "duplicate item"],
        "paste": ["paste file", "insert item"],
        "cut": ["cut file", "move item"],
        "select all": ["select everything", "highlight all"],
        "view": ["change view", "list view"],
        "sort": ["sort files", "order items"],
        "search": ["find file", "search folder"],
        "address": ["go to address", "navigate to"],
        "back": ["go back", "previous folder"],
        "forward": ["go forward", "next folder"],
        "up": ["go up", "parent folder"]
    }
    
    if not fuzzy_match(command, ["file explorer", "files", "folder"]):
        return False
    
    # Open File Explorer if not already open
    try:
        os.system("start explorer")
        time.sleep(2)  # Wait for File Explorer to open
    except:
        speak("Couldn't open File Explorer", 'high')
        return True
    
    for action, keywords in explorer_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "new folder":
                    keyboard.press_and_release('ctrl+shift+n')
                    speak("New folder created", 'normal')
                elif action == "rename":
                    keyboard.press_and_release('f2')
                    speak("Ready to rename selected item", 'normal')
                elif action == "delete":
                    keyboard.press_and_release('del')
                    speak("Item moved to recycle bin", 'normal')
                elif action == "properties":
                    keyboard.press_and_release('alt+enter')
                    speak("Showing properties", 'normal')
                elif action == "open":
                    keyboard.press_and_release('enter')
                    speak("Opening selected item", 'normal')
                elif action == "copy":
                    keyboard.press_and_release('ctrl+c')
                    speak("Item copied", 'normal')
                elif action == "paste":
                    keyboard.press_and_release('ctrl+v')
                    speak("Item pasted", 'normal')
                elif action == "cut":
                    keyboard.press_and_release('ctrl+x')
                    speak("Item cut", 'normal')
                elif action == "select all":
                    keyboard.press_and_release('ctrl+a')
                    speak("All items selected", 'normal')
                elif action == "view":
                    if "details" in command:
                        keyboard.press_and_release('ctrl+shift+6')
                        speak("Details view enabled", 'normal')
                    elif "list" in command:
                        keyboard.press_and_release('ctrl+shift+2')
                        speak("List view enabled", 'normal')
                    elif "tiles" in command:
                        keyboard.press_and_release('ctrl+shift+3')
                        speak("Tiles view enabled", 'normal')
                    elif "icons" in command:
                        keyboard.press_and_release('ctrl+shift+1')
                        speak("Icons view enabled", 'normal')
                    else:
                        keyboard.press_and_release('ctrl+shift+5')
                        speak("View changed", 'normal')
                elif action == "sort":
                    if "name" in command:
                        keyboard.press_and_release('ctrl+shift+1')
                        speak("Sorted by name", 'normal')
                    elif "date" in command:
                        keyboard.press_and_release('ctrl+shift+4')
                        speak("Sorted by date", 'normal')
                    elif "size" in command:
                        keyboard.press_and_release('ctrl+shift+2')
                        speak("Sorted by size", 'normal')
                    elif "type" in command:
                        keyboard.press_and_release('ctrl+shift+3')
                        speak("Sorted by type", 'normal')
                    else:
                        keyboard.press_and_release('ctrl+shift+1')
                        speak("Items sorted", 'normal')
                elif action == "search":
                    keyboard.press_and_release('ctrl+f')
                    time.sleep(0.5)
                    speak("What would you like to search for?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "address":
                    keyboard.press_and_release('alt+d')
                    time.sleep(0.5)
                    speak("Where would you like to go?", 'normal')
                    path = listen()
                    if path:
                        type_text(path)
                        keyboard.press_and_release('enter')
                        speak(f"Navigating to {path}", 'normal')
                elif action == "back":
                    keyboard.press_and_release('alt+left')
                    speak("Going back", 'normal')
                elif action == "forward":
                    keyboard.press_and_release('alt+right')
                    speak("Going forward", 'normal')
                elif action == "up":
                    keyboard.press_and_release('alt+up')
                    speak("Going up one level", 'normal')
                return True
            except Exception as e:
                print(f"File Explorer control error: {e}")
                speak("Couldn't complete that File Explorer action", 'high')
                return True
    
    # Handle navigation to specific folders
    if fuzzy_match(command, ["go to", "navigate to"]):
        folder = command.replace("go to", "").replace("navigate to", "").strip()
        if folder:
            try:
                keyboard.press_and_release('alt+d')
                time.sleep(0.5)
                type_text(folder)
                keyboard.press_and_release('enter')
                speak(f"Navigating to {folder}", 'normal')
                return True
            except:
                speak(f"Couldn't navigate to {folder}", 'high')
                return True
    
    # Default File Explorer opening
    if fuzzy_match(command, ["file explorer", "open files"]):
        speak("File Explorer is already open", 'normal')
        return True
    
    return False

def handle_chrome_commands(command):
    """Enhanced Chrome browser control"""
    chrome_actions = {
        "new tab": ["open tab", "new tab"],
        "close tab": ["close tab", "shut tab"],
        "reopen tab": ["reopen tab", "restore tab"],
        "next tab": ["next tab", "switch right"],
        "previous tab": ["previous tab", "switch left"],
        "bookmark": ["add bookmark", "save page"],
        "history": ["view history", "browsing history"],
        "downloads": ["show downloads", "my downloads"],
        "extensions": ["view extensions", "my addons"],
        "incognito": ["private mode", "incognito mode"],
        "zoom in": ["zoom in", "increase zoom"],
        "zoom out": ["zoom out", "decrease zoom"],
        "reset zoom": ["normal zoom", "default zoom"],
        "find": ["find on page", "search page"],
        "refresh": ["reload page", "refresh page"],
        "stop": ["stop loading", "cancel load"],
        "home": ["go home", "home page"],
        "developer": ["developer tools", "inspect element"],
        "task manager": ["browser tasks", "chrome processes"],
        "clear data": ["clear cache", "delete cookies"]
    }
    
    if not fuzzy_match(command, ["chrome", "google chrome", "browser"]):
        return False
    
    # Focus on Chrome window
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("%")
        win32gui.SetForegroundWindow(win32gui.FindWindow(None, "Google Chrome"))
        time.sleep(1)
    except:
        speak("Couldn't focus on Chrome", 'high')
        return True
    
    for action, keywords in chrome_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "new tab":
                    keyboard.press_and_release('ctrl+t')
                    speak("New tab opened", 'normal')
                elif action == "close tab":
                    keyboard.press_and_release('ctrl+w')
                    speak("Tab closed", 'normal')
                elif action == "reopen tab":
                    keyboard.press_and_release('ctrl+shift+t')
                    speak("Last closed tab reopened", 'normal')
                elif action == "next tab":
                    keyboard.press_and_release('ctrl+tab')
                    speak("Switched to next tab", 'normal')
                elif action == "previous tab":
                    keyboard.press_and_release('ctrl+shift+tab')
                    speak("Switched to previous tab", 'normal')
                elif action == "bookmark":
                    keyboard.press_and_release('ctrl+d')
                    speak("Bookmark dialog opened", 'normal')
                elif action == "history":
                    keyboard.press_and_release('ctrl+h')
                    speak("Browsing history shown", 'normal')
                elif action == "downloads":
                    keyboard.press_and_release('ctrl+j')
                    speak("Downloads shown", 'normal')
                elif action == "extensions":
                    webbrowser.open("chrome://extensions")
                    speak("Extensions page opened", 'normal')
                elif action == "incognito":
                    keyboard.press_and_release('ctrl+shift+n')
                    speak("Incognito window opened", 'normal')
                elif action == "zoom in":
                    keyboard.press_and_release('ctrl+plus')
                    speak("Zoomed in", 'normal')
                elif action == "zoom out":
                    keyboard.press_and_release('ctrl+minus')
                    speak("Zoomed out", 'normal')
                elif action == "reset zoom":
                    keyboard.press_and_release('ctrl+0')
                    speak("Zoom reset", 'normal')
                elif action == "find":
                    keyboard.press_and_release('ctrl+f')
                    time.sleep(0.5)
                    speak("What would you like to find?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "refresh":
                    keyboard.press_and_release('f5')
                    speak("Page refreshed", 'normal')
                elif action == "stop":
                    keyboard.press_and_release('esc')
                    speak("Loading stopped", 'normal')
                elif action == "home":
                    keyboard.press_and_release('alt+home')
                    speak("Navigated to home page", 'normal')
                elif action == "developer":
                    keyboard.press_and_release('ctrl+shift+i')
                    speak("Developer tools opened", 'normal')
                elif action == "task manager":
                    keyboard.press_and_release('shift+esc')
                    speak("Browser task manager opened", 'normal')
                elif action == "clear data":
                    webbrowser.open("chrome://settings/clearBrowserData")
                    speak("Clear browsing data dialog opened", 'normal')
                return True
            except Exception as e:
                print(f"Chrome control error: {e}")
                speak("Couldn't complete that Chrome action", 'high')
                return True
    
    # Handle Chrome searches
    if fuzzy_match(command, ["search for", "look up"]):
        query = command.replace("search for", "").replace("look up", "").strip()
        if not query:
            speak("What would you like to search for?", 'normal')
            query = listen()
        if query:
            keyboard.press_and_release('ctrl+l')  # Focus address bar
            time.sleep(0.5)
            type_text(query)
            keyboard.press_and_release('enter')
            speak(f"Searching for {query}", 'normal')
            return True
    
    # Handle URL navigation
    if fuzzy_match(command, ["go to", "navigate to"]):
        url = command.replace("go to", "").replace("navigate to", "").strip()
        if not url.startswith("http"):
            url = "https://" + url
        if "." not in url:
            url += ".com"
        
        try:
            keyboard.press_and_release('ctrl+l')  # Focus address bar
            time.sleep(0.5)
            type_text(url)
            keyboard.press_and_release('enter')
            speak(f"Navigating to {url}", 'normal')
            return True
        except:
            speak(f"Couldn't navigate to {url}", 'high')
            return True
    
    # Default Chrome opening
    if fuzzy_match(command, ["chrome", "open chrome"]):
        webbrowser.get("chrome").open("https://www.google.com")
        speak("Chrome opened", 'normal')
        return True
    
    return False

def handle_google_commands(command):
    """Enhanced Google search operations"""
    google_actions = {
        "search": ["google search", "web search"],
        "images": ["image search", "find pictures"],
        "videos": ["video search", "find videos"],
        "news": ["news search", "find articles"],
        "maps": ["map search", "find location"],
        "flights": ["flight search", "find flights"],
        "hotels": ["hotel search", "find hotels"],
        "shopping": ["product search", "find items"],
        "books": ["book search", "find books"],
        "scholar": ["academic search", "find papers"],
        "finance": ["stock search", "find stocks"],
        "translate": ["language translation", "translate text"]
    }
    
    if not fuzzy_match(command, ["google", "search"]):
        return False
    
    for action, keywords in google_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "search":
                    query = command.replace("search", "").replace("google", "").strip()
                    if not query:
                        speak("What would you like to search for?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?q={query}")
                        speak(f"Searching Google for {query}", 'normal')
                elif action == "images":
                    query = command.replace("images", "").replace("pictures", "").strip()
                    if not query:
                        speak("What images would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=isch&q={query}")
                        speak(f"Searching images for {query}", 'normal')
                elif action == "videos":
                    query = command.replace("videos", "").replace("video", "").strip()
                    if not query:
                        speak("What videos would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=vid&q={query}")
                        speak(f"Searching videos for {query}", 'normal')
                elif action == "news":
                    query = command.replace("news", "").replace("articles", "").strip()
                    if not query:
                        speak("What news would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=nws&q={query}")
                        speak(f"Searching news for {query}", 'normal')
                elif action == "maps":
                    query = command.replace("maps", "").replace("location", "").strip()
                    if not query:
                        speak("What location would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/maps/search/{query}")
                        speak(f"Searching maps for {query}", 'normal')
                elif action == "flights":
                    query = command.replace("flights", "").replace("flight", "").strip()
                    if not query:
                        speak("What flight would you like to search for?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=flm&q={query}")
                        speak(f"Searching flights for {query}", 'normal')
                elif action == "hotels":
                    query = command.replace("hotels", "").replace("hotel", "").strip()
                    if not query:
                        speak("What hotel would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=lod&q={query}")
                        speak(f"Searching hotels for {query}", 'normal')
                elif action == "shopping":
                    query = command.replace("shopping", "").replace("products", "").strip()
                    if not query:
                        speak("What product would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=shop&q={query}")
                        speak(f"Searching products for {query}", 'normal')
                elif action == "books":
                    query = command.replace("books", "").replace("book", "").strip()
                    if not query:
                        speak("What book would you like to find?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/search?tbm=bks&q={query}")
                        speak(f"Searching books for {query}", 'normal')
                elif action == "scholar":
                    query = command.replace("scholar", "").replace("papers", "").strip()
                    if not query:
                        speak("What would you like to search in Google Scholar?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://scholar.google.com/scholar?q={query}")
                        speak(f"Searching scholar for {query}", 'normal')
                elif action == "finance":
                    query = command.replace("finance", "").replace("stocks", "").strip()
                    if not query:
                        speak("What stock would you like to check?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://www.google.com/finance/quote/{query}")
                        speak(f"Searching finance for {query}", 'normal')
                elif action == "translate":
                    query = command.replace("translate", "").strip()
                    if not query:
                        speak("What would you like to translate?", 'normal')
                        query = listen()
                    if query:
                        webbrowser.open(f"https://translate.google.com/?text={query}")
                        speak(f"Translating: {query}", 'normal')
                return True
            except Exception as e:
                print(f"Google search error: {e}")
                speak("Couldn't complete that Google search", 'high')
                return True
    
    # Default Google search
    if fuzzy_match(command, ["google", "search google"]):
        webbrowser.open("https://www.google.com")
        speak("Google opened", 'normal')
        return True
    
    return False

def handle_cmd_commands(command):
    """Enhanced Command Prompt operations"""
    cmd_actions = {
        "run": ["run command", "execute"],
        "admin": ["admin prompt", "elevated"],
        "clear": ["clear screen", "reset"],
        "copy": ["copy text", "select all"],
        "paste": ["paste text", "insert"],
        "history": ["command history", "previous commands"],
        "dir": ["list files", "show directory"],
        "cd": ["change directory", "navigate to"],
        "ipconfig": ["network info", "ip address"],
        "ping": ["test connection", "ping server"],
        "tracert": ["trace route", "network path"],
        "netstat": ["network stats", "connections"],
        "tasklist": ["running processes", "show tasks"],
        "systeminfo": ["system information", "computer info"],
        "chkdsk": ["check disk", "scan drive"],
        "sfc": ["system scan", "file check"],
        "diskpart": ["disk utility", "partition tool"],
        "gpupdate": ["group policy", "update settings"]
    }
    
    if not fuzzy_match(command, ["command prompt", "cmd", "terminal"]):
        return False
    
    # Open Command Prompt if not already open
    try:
        os.system("start cmd")
        time.sleep(2)  # Wait for CMD to open
    except:
        speak("Couldn't open Command Prompt", 'high')
        return True
    
    for action, keywords in cmd_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "run":
                    speak("What command would you like to run?", 'normal')
                    cmd = listen()
                    if cmd:
                        type_text(cmd)
                        keyboard.press_and_release('enter')
                        speak("Command executed", 'normal')
                elif action == "admin":
                    os.system("start cmd /k echo Running as administrator")
                    speak("Administrator command prompt opened", 'normal')
                elif action == "clear":
                    keyboard.press_and_release('ctrl+l')
                    speak("Screen cleared", 'normal')
                elif action == "copy":
                    keyboard.press_and_release('ctrl+m')
                    time.sleep(0.5)
                    keyboard.press_and_release('ctrl+c')
                    speak("Text copied", 'normal')
                elif action == "paste":
                    keyboard.press_and_release('ctrl+v')
                    speak("Text pasted", 'normal')
                elif action == "history":
                    keyboard.press_and_release('f7')
                    speak("Command history shown", 'normal')
                elif action == "dir":
                    type_text("dir")
                    keyboard.press_and_release('enter')
                    speak("Directory listing shown", 'normal')
                elif action == "cd":
                    speak("Which directory would you like to go to?", 'normal')
                    path = listen()
                    if path:
                        type_text(f"cd {path}")
                        keyboard.press_and_release('enter')
                        speak(f"Changed to {path} directory", 'normal')
                elif action == "ipconfig":
                    type_text("ipconfig /all")
                    keyboard.press_and_release('enter')
                    speak("Network information shown", 'normal')
                elif action == "ping":
                    speak("What would you like to ping?", 'normal')
                    target = listen()
                    if target:
                        type_text(f"ping {target}")
                        keyboard.press_and_release('enter')
                        speak(f"Pinging {target}", 'normal')
                elif action == "tracert":
                    speak("What would you like to trace?", 'normal')
                    target = listen()
                    if target:
                        type_text(f"tracert {target}")
                        keyboard.press_and_release('enter')
                        speak(f"Tracing route to {target}", 'normal')
                elif action == "netstat":
                    type_text("netstat -ano")
                    keyboard.press_and_release('enter')
                    speak("Network statistics shown", 'normal')
                elif action == "tasklist":
                    type_text("tasklist")
                    keyboard.press_and_release('enter')
                    speak("Running processes shown", 'normal')
                elif action == "systeminfo":
                    type_text("systeminfo")
                    keyboard.press_and_release('enter')
                    speak("System information shown", 'normal')
                elif action == "chkdsk":
                    speak("Which drive would you like to check?", 'normal')
                    drive = listen()
                    if drive:
                        type_text(f"chkdsk {drive}:")
                        keyboard.press_and_release('enter')
                        speak(f"Checking {drive} drive", 'normal')
                elif action == "sfc":
                    type_text("sfc /scannow")
                    keyboard.press_and_release('enter')
                    speak("System file check started", 'normal')
                elif action == "diskpart":
                    type_text("diskpart")
                    keyboard.press_and_release('enter')
                    speak("Diskpart utility opened", 'normal')
                elif action == "gpupdate":
                    type_text("gpupdate /force")
                    keyboard.press_and_release('enter')
                    speak("Group policy update forced", 'normal')
                return True
            except Exception as e:
                print(f"Command Prompt error: {e}")
                speak("Couldn't complete that Command Prompt action", 'high')
                return True
    
    # Default Command Prompt opening
    if fuzzy_match(command, ["command prompt", "open cmd"]):
        speak("Command Prompt is already open", 'normal')
        return True
    
    return False

# ===================== SYSTEM COMMAND HANDLERS =====================

def handle_media_control_commands(command):
    """Handle media playback controls"""
    media_actions = {
        "play": ["play", "start", "resume"],
        "pause": ["pause", "stop"],
        "next": ["next", "skip"],
        "previous": ["previous", "back"],
        "volume up": ["volume up", "louder"],
        "volume down": ["volume down", "quieter"],
        "mute": ["mute", "silence"],
        "fullscreen": ["fullscreen", "maximize"],
        "theater mode": ["theater mode", "cinema view"],
        "miniplayer": ["miniplayer", "small window"],
        "captions": ["captions", "subtitles"],
        "speed up": ["speed up", "faster"],
        "slow down": ["slow down", "slower"],
        "normal speed": ["normal speed", "reset speed"]
    }
    
    global media_playing, current_volume
    
    for action, keywords in media_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "play":
                    keyboard.press_and_release('k')  # YouTube play/pause
                    media_playing = True
                    speak("Playback started", 'normal')
                elif action == "pause":
                    keyboard.press_and_release('k')  # YouTube play/pause
                    media_playing = False
                    speak("Playback paused", 'normal')
                elif action == "next":
                    keyboard.press_and_release('shift+n')  # YouTube next
                    speak("Skipped to next", 'normal')
                elif action == "previous":
                    keyboard.press_and_release('shift+p')  # YouTube previous
                    speak("Went back to previous", 'normal')
                elif action == "volume up":
                    current_volume = min(100, current_volume + 10)
                    pyautogui.press('volumeup')
                    speak(f"Volume increased to {current_volume}%", 'normal')
                elif action == "volume down":
                    current_volume = max(0, current_volume - 10)
                    pyautogui.press('volumedown')
                    speak(f"Volume decreased to {current_volume}%", 'normal')
                elif action == "mute":
                    pyautogui.press('volumemute')
                    speak("Volume muted", 'normal')
                elif action == "fullscreen":
                    keyboard.press_and_release('f')
                    speak("Toggled fullscreen", 'normal')
                elif action == "theater mode":
                    keyboard.press_and_release('t')
                    speak("Toggled theater mode", 'normal')
                elif action == "miniplayer":
                    keyboard.press_and_release('i')
                    speak("Toggled miniplayer", 'normal')
                elif action == "captions":
                    keyboard.press_and_release('c')
                    speak("Toggled captions", 'normal')
                elif action == "speed up":
                    keyboard.press_and_release('shift+>')
                    speak("Playback speed increased", 'normal')
                elif action == "slow down":
                    keyboard.press_and_release('shift+<')
                    speak("Playback speed decreased", 'normal')
                elif action == "normal speed":
                    keyboard.press_and_release('shift+/')
                    speak("Playback speed reset", 'normal')
                return True
            except Exception as e:
                print(f"Media control error: {e}")
                speak("Couldn't complete that media action", 'high')
                return True
    return False

def handle_system_commands(command):
    """Handle system-level commands"""
    system_actions = {
        "shutdown": ["shutdown", "turn off"],
        "restart": ["restart", "reboot"],
        "sleep": ["sleep", "suspend"],
        "lock": ["lock", "secure"],
        "logout": ["logout", "sign out"],
        "task manager": ["task manager", "processes"],
        "screenshot": ["screenshot", "capture"],
        "clipboard": ["clipboard", "pasteboard"],
        "brightness up": ["brightness up", "brighter"],
        "brightness down": ["brightness down", "dimmer"],
        "wifi": ["wifi", "wireless"],
        "bluetooth": ["bluetooth", "wireless devices"],
        "airplane mode": ["airplane mode", "flight mode"],
        "battery": ["battery", "power"],
        "display": ["display", "monitor"],
        "sound": ["sound", "audio"],
        "notifications": ["notifications", "alerts"],
        "calendar": ["calendar", "schedule"],
        "clock": ["clock", "time"],
        "calculator": ["calculator", "calc"],
        "camera": ["camera", "webcam"],
        "microphone": ["microphone", "mic"],
        "printer": ["printer", "printing"]
    }
    
    for action, keywords in system_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "shutdown":
                    if confirm_action("shutdown the computer"):
                        os.system("shutdown /s /t 1")
                        speak("Shutting down", 'high')
                elif action == "restart":
                    if confirm_action("restart the computer"):
                        os.system("shutdown /r /t 1")
                        speak("Restarting", 'high')
                elif action == "sleep":
                    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                    speak("Putting computer to sleep", 'high')
                elif action == "lock":
                    os.system("rundll32.exe user32.dll,LockWorkStation")
                    speak("Locking computer", 'high')
                elif action == "logout":
                    os.system("shutdown /l")
                    speak("Logging out", 'high')
                elif action == "task manager":
                    os.system("start taskmgr")
                    speak("Task manager opened", 'normal')
                elif action == "screenshot":
                    pyautogui.screenshot("screenshot.png")
                    speak("Screenshot taken and saved", 'normal')
                elif action == "clipboard":
                    clipboard_content = clipboard.paste()
                    if clipboard_content:
                        speak(f"Clipboard contains: {clipboard_content}", 'normal')
                    else:
                        speak("Clipboard is empty", 'normal')
                elif action == "brightness up":
                    current = sbc.get_brightness()[0]
                    new = min(100, current + 10)
                    sbc.set_brightness(new)
                    speak(f"Brightness increased to {new}%", 'normal')
                elif action == "brightness down":
                    current = sbc.get_brightness()[0]
                    new = max(0, current - 10)
                    sbc.set_brightness(new)
                    speak(f"Brightness decreased to {new}%", 'normal')
                elif action == "wifi":
                    os.system("start ms-availablenetworks:")
                    speak("Wi-Fi networks shown", 'normal')
                elif action == "bluetooth":
                    os.system("start ms-settings:bluetooth")
                    speak("Bluetooth settings opened", 'normal')
                elif action == "airplane mode":
                    os.system("start ms-settings:network-airplanemode")
                    speak("Airplane mode settings opened", 'normal')
                elif action == "battery":
                    battery = psutil.sensors_battery()
                    percent = battery.percent
                    plugged = "plugged in" if battery.power_plugged else "not plugged in"
                    speak(f"Battery is at {percent}% and {plugged}", 'normal')
                elif action == "display":
                    os.system("start ms-settings:display")
                    speak("Display settings opened", 'normal')
                elif action == "sound":
                    os.system("start ms-settings:sound")
                    speak("Sound settings opened", 'normal')
                elif action == "notifications":
                    os.system("start ms-settings:notifications")
                    speak("Notification settings opened", 'normal')
                elif action == "calendar":
                    os.system("start outlookcal:")
                    speak("Calendar opened", 'normal')
                elif action == "clock":
                    now = datetime.datetime.now().strftime("%I:%M %p")
                    speak(f"The current time is {now}", 'normal')
                
                elif action == "calculator":
                    os.system("start calc")
                    speak("Calculator opened", 'normal')
                elif action == "camera":
                    os.system("start microsoft.windows.camera:")
                    speak("Camera app opened", 'normal')
                elif action == "microphone":
                    os.system("start ms-settings:privacy-microphone")
                    speak("Microphone settings opened", 'normal')
                elif action == "printer":
                    os.system("start ms-settings:printers")
                    speak("Printer settings opened", 'normal')
                return True
            except Exception as e:
                print(f"System command error: {e}")
                speak("Couldn't complete that system action", 'high')
                return True
    return False

def handle_app_commands(command):
    """Handle application launch commands"""
    apps = {
        "word": ["word", "microsoft word"],
        "excel": ["excel", "microsoft excel"],
        "powerpoint": ["powerpoint", "microsoft powerpoint"],
        "outlook": ["outlook", "microsoft outlook"],
        "onenote": ["onenote", "microsoft onenote"],
        "teams": ["teams", "microsoft teams"],
        "edge": ["edge", "microsoft edge"],
        "photoshop": ["photoshop", "adobe photoshop"],
        "premiere": ["premiere", "adobe premiere"],
        "illustrator": ["illustrator", "adobe illustrator"],
        "spotify": ["spotify", "music app"],
        "netflix": ["netflix", "streaming"],
        "zoom": ["zoom", "video call"],
        "discord": ["discord", "chat app"],
        "slack": ["slack", "team chat"],
        "vscode": ["vscode", "visual studio code"],
        "pycharm": ["pycharm", "python ide"],
        "eclipse": ["eclipse", "java ide"],
        "blender": ["blender", "3d modeling"],
        "steam": ["steam", "game platform"],
        "epic": ["epic games", "epic launcher"]
    }
    
    for app, keywords in apps.items():
        if fuzzy_match(command, keywords):
            try:
                if app == "word":
                    os.system("start winword")
                    speak("Microsoft Word opened", 'normal')
                elif app == "excel":
                    os.system("start excel")
                    speak("Microsoft Excel opened", 'normal')
                elif app == "powerpoint":
                    os.system("start powerpnt")
                    speak("Microsoft PowerPoint opened", 'normal')
                elif app == "outlook":
                    os.system("start outlook")
                    speak("Microsoft Outlook opened", 'normal')
                elif app == "onenote":
                    os.system("start onenote")
                    speak("Microsoft OneNote opened", 'normal')
                elif app == "teams":
                    os.system("start msteams")
                    speak("Microsoft Teams opened", 'normal')
                elif app == "edge":
                    os.system("start msedge")
                    speak("Microsoft Edge opened", 'normal')
                elif app == "photoshop":
                    os.system("start photoshop")
                    speak("Adobe Photoshop opened", 'normal')
                elif app == "premiere":
                    os.system("start premiere")
                    speak("Adobe Premiere opened", 'normal')
                elif app == "illustrator":
                    os.system("start illustrator")
                    speak("Adobe Illustrator opened", 'normal')
                elif app == "spotify":
                    os.system("start spotify")
                    speak("Spotify opened", 'normal')
                elif app == "netflix":
                    webbrowser.open("https://www.netflix.com")
                    speak("Netflix opened", 'normal')
                elif app == "zoom":
                    os.system("start zoom")
                    speak("Zoom opened", 'normal')
                elif app == "discord":
                    os.system("start discord")
                    speak("Discord opened", 'normal')
                elif app == "slack":
                    os.system("start slack")
                    speak("Slack opened", 'normal')
                elif app == "vscode":
                    os.system("start code")
                    speak("VS Code opened", 'normal')
                elif app == "pycharm":
                    os.system("start pycharm")
                    speak("PyCharm opened", 'normal')
                elif app == "eclipse":
                    os.system("start eclipse")
                    speak("Eclipse opened", 'normal')
                elif app == "blender":
                    os.system("start blender")
                    speak("Blender opened", 'normal')
                elif app == "steam":
                    os.system("start steam")
                    speak("Steam opened", 'normal')
                elif app == "epic":
                    os.system("start com.epicgames.launcher")
                    speak("Epic Games Launcher opened", 'normal')
                return True
            except Exception as e:
                print(f"App launch error: {e}")
                speak(f"Couldn't open {app}", 'high')
                return True
    return False

def handle_web_search_commands(command):
    """Handle web search commands"""
    search_engines = {
        "google": ["google", "search google"],
        "bing": ["bing", "search bing"],
        "yahoo": ["yahoo", "search yahoo"],
        "duckduckgo": ["duckduckgo", "search duckduckgo"],
        "youtube": ["youtube", "search youtube"],
        "wikipedia": ["wikipedia", "search wikipedia"],
        "amazon": ["amazon", "search amazon"],
        "ebay": ["ebay", "search ebay"],
        "twitter": ["twitter", "search twitter"],
        "facebook": ["facebook", "search facebook"],
        "instagram": ["instagram", "search instagram"],
        "reddit": ["reddit", "search reddit"],
        "linkedin": ["linkedin", "search linkedin"],
        "pinterest": ["pinterest", "search pinterest"],
        "imdb": ["imdb", "search imdb"],
        "rottentomatoes": ["rottentomatoes", "search rottentomatoes"],
        "github": ["github", "search github"],
        "stackoverflow": ["stackoverflow", "search stackoverflow"]
    }
    
    for engine, keywords in search_engines.items():
        if fuzzy_match(command, keywords):
            query = command.replace(engine, "").replace("search", "").strip()
            if not query:
                speak(f"What would you like to search on {engine}?", 'normal')
                query = listen()
            if query:
                if engine == "google":
                    webbrowser.open(f"https://www.google.com/search?q={query}")
                elif engine == "bing":
                    webbrowser.open(f"https://www.bing.com/search?q={query}")
                elif engine == "yahoo":
                    webbrowser.open(f"https://search.yahoo.com/search?p={query}")
                elif engine == "duckduckgo":
                    webbrowser.open(f"https://duckduckgo.com/?q={query}")
                elif engine == "youtube":
                    webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
                elif engine == "wikipedia":
                    webbrowser.open(f"https://en.wikipedia.org/wiki/Special:Search?search={query}")
                elif engine == "amazon":
                    webbrowser.open(f"https://www.amazon.com/s?k={query}")
                elif engine == "ebay":
                    webbrowser.open(f"https://www.ebay.com/sch/i.html?_nkw={query}")
                elif engine == "twitter":
                    webbrowser.open(f"https://twitter.com/search?q={query}")
                elif engine == "facebook":
                    webbrowser.open(f"https://www.facebook.com/search/top/?q={query}")
                elif engine == "instagram":
                    webbrowser.open(f"https://www.instagram.com/explore/tags/{query}/")
                elif engine == "reddit":
                    webbrowser.open(f"https://www.reddit.com/search/?q={query}")
                elif engine == "linkedin":
                    webbrowser.open(f"https://www.linkedin.com/search/results/all/?keywords={query}")
                elif engine == "pinterest":
                    webbrowser.open(f"https://www.pinterest.com/search/pins/?q={query}")
                elif engine == "imdb":
                    webbrowser.open(f"https://www.imdb.com/find?q={query}")
                elif engine == "rottentomatoes":
                    webbrowser.open(f"https://www.rottentomatoes.com/search?search={query}")
                elif engine == "github":
                    webbrowser.open(f"https://github.com/search?q={query}")
                elif engine == "stackoverflow":
                    webbrowser.open(f"https://stackoverflow.com/search?q={query}")
                speak(f"Searching {engine} for {query}", 'normal')
                return True
    return False

def handle_weather_commands(command):
    """Handle weather-related commands"""
    if not fuzzy_match(command, ["weather", "forecast", "temperature"]):
        return False
    
    try:
        if "current" in command or "now" in command:
            speak("For which city?", 'normal')
            city = listen()
            if city:
                mgr = owm.weather_manager()
                observation = mgr.weather_at_place(city)
                w = observation.weather
                temp = w.temperature('celsius')['temp']
                status = w.detailed_status
                speak(f"Current weather in {city}: {status}, temperature {temp}°C", 'normal')
        elif "forecast" in command or "tomorrow" in command:
            speak("For which city?", 'normal')
            city = listen()
            if city:
                mgr = owm.weather_manager()
                forecast = mgr.forecast_at_place(city, '3h')
                tomorrow = pyowm.utils.timestamps.tomorrow()
                weather = forecast.get_weather_at(tomorrow)
                temp = weather.temperature('celsius')['temp']
                status = weather.detailed_status
                speak(f"Tomorrow's forecast for {city}: {status}, temperature around {temp}°C", 'normal')
        else:
            speak("For which city?", 'normal')
            city = listen()
            if city:
                mgr = owm.weather_manager()
                observation = mgr.weather_at_place(city)
                w = observation.weather
                temp = w.temperature('celsius')['temp']
                status = w.detailed_status
                humidity = w.humidity
                wind = w.wind()['speed']
                speak(f"Weather in {city}: {status}, temperature {temp}°C, humidity {humidity}%, wind speed {wind} m/s", 'normal')
        return True
    except Exception as e:
        print(f"Weather error: {e}")
        speak("Couldn't get weather information", 'high')
        return True

def handle_news_commands(command):
    """Handle news-related commands"""
    if not fuzzy_match(command, ["news", "headlines", "updates"]):
        return False
    
    try:
        if "business" in command:
            news = newsapi.get_top_headlines(category='business', language='en')
            speak("Here are the top business headlines", 'normal')
        elif "technology" in command or "tech" in command:
            news = newsapi.get_top_headlines(category='technology', language='en')
            speak("Here are the top technology headlines", 'normal')
        elif "sports" in command:
            news = newsapi.get_top_headlines(category='sports', language='en')
            speak("Here are the top sports headlines", 'normal')
        elif "health" in command:
            news = newsapi.get_top_headlines(category='health', language='en')
            speak("Here are the top health headlines", 'normal')
        elif "science" in command:
            news = newsapi.get_top_headlines(category='science', language='en')
            speak("Here are the top science headlines", 'normal')
        elif "entertainment" in command:
            news = newsapi.get_top_headlines(category='entertainment', language='en')
            speak("Here are the top entertainment headlines", 'normal')
        else:
            news = newsapi.get_top_headlines(language='en')
            speak("Here are the top headlines", 'normal')
        
        articles = news['articles']
        for i, article in enumerate(articles[:5]):
            speak(f"{i+1}. {article['title']}", 'normal')
            if i < 4:  # Pause between headlines except last one
                time.sleep(1)
        return True
    except Exception as e:
        print(f"News error: {e}")
        speak("Couldn't get news updates", 'high')
        return True

def handle_wikipedia_commands(command):
    """Handle Wikipedia search commands"""
    if not fuzzy_match(command, ["wikipedia", "wiki"]):
        return False
    
    query = command.replace("wikipedia", "").replace("wiki", "").strip()
    if not query:
        speak("What would you like to search on Wikipedia?", 'normal')
        query = listen()
    if query:
        try:
            summary = wikipedia.summary(query, sentences=2)
            speak(f"According to Wikipedia: {summary}", 'normal')
        except wikipedia.exceptions.DisambiguationError as e:
            speak(f"Multiple results found. Can you be more specific? Some options are: {', '.join(e.options[:5])}", 'normal')
        except wikipedia.exceptions.PageError:
            speak("No Wikipedia page found for that query", 'normal')
        except Exception as e:
            print(f"Wikipedia error: {e}")
            speak("Couldn't get Wikipedia information", 'high')
    return True

def handle_email_commands(command):
    """Handle email-related commands"""
    email_actions = {
        "send": ["send email", "compose"],
        "read": ["read emails", "check inbox"],
        "reply": ["reply email", "respond"],
        "forward": ["forward email", "send copy"],
        "delete": ["delete email", "remove message"],
        "search": ["search emails", "find message"],
        "attachment": ["add attachment", "include file"],
        "draft": ["save draft", "temporary save"],
        "contacts": ["view contacts", "address book"],
        "settings": ["email settings", "account settings"]
    }
    
    if not fuzzy_match(command, ["email", "gmail", "outlook"]):
        return False
    
    for action, keywords in email_actions.items():
        if fuzzy_match(command, keywords):
            try:
                if action == "send":
                    speak("What's the subject of the email?", 'normal')
                    subject = listen()
                    speak("What should the message say?", 'normal')
                    body = listen()
                    speak("Who should I send it to?", 'normal')
                    to = listen()
                    
                    msg = EmailMessage()
                    msg.set_content(body)
                    msg['Subject'] = subject
                    msg['From'] = "your_email@example.com"
                    msg['To'] = to
                    
                    with smtplib.SMTP('smtp.gmail.com', 587) as server:
                        server.starttls()
                        server.login("your_email@example.com", "your_password")
                        server.send_message(msg)
                    
                    speak("Email sent successfully", 'normal')
                elif action == "read":
                    webbrowser.open("https://mail.google.com")
                    speak("Opening your inbox", 'normal')
                elif action == "reply":
                    speak("What should the reply say?", 'normal')
                    reply = listen()
                    if reply:
                        keyboard.press_and_release('r')  # Gmail reply
                        time.sleep(1)
                        type_text(reply)
                        keyboard.press_and_release('ctrl+enter')
                        speak("Reply sent", 'normal')
                elif action == "forward":
                    keyboard.press_and_release('f')  # Gmail forward
                    speak("Who should I forward this to?", 'normal')
                    to = listen()
                    if to:
                        keyboard.press_and_release('tab')
                        type_text(to)
                        keyboard.press_and_release('ctrl+enter')
                        speak("Email forwarded", 'normal')
                elif action == "delete":
                    keyboard.press_and_release('#')  # Gmail delete
                    speak("Email deleted", 'normal')
                elif action == "search":
                    keyboard.press_and_release('/')  # Gmail search
                    speak("What would you like to search for?", 'normal')
                    search_term = listen()
                    if search_term:
                        type_text(search_term)
                        keyboard.press_and_release('enter')
                        speak(f"Searching for {search_term}", 'normal')
                elif action == "attachment":
                    keyboard.press_and_release('ctrl+shift+a')  # Gmail attach
                    speak("Please select the file to attach", 'normal')
                elif action == "draft":
                    keyboard.press_and_release('ctrl+s')  # Gmail save draft
                    speak("Draft saved", 'normal')
                elif action == "contacts":
                    webbrowser.open("https://contacts.google.com")
                    speak("Opening contacts", 'normal')
                elif action == "settings":
                    keyboard.press_and_release('shift+?')  # Gmail settings
                    time.sleep(0.5)
                    keyboard.press_and_release('s')
                    speak("Settings opened", 'normal')
                return True
            except Exception as e:
                print(f"Email error: {e}")
                speak("Couldn't complete that email action", 'high')
                return True
    
    # Default email opening
    if fuzzy_match(command, ["email", "open email"]):
        webbrowser.open("https://mail.google.com")
        speak("Opening your email", 'normal')
        return True
    
    return False

def handle_math_commands(command):
    """Handle mathematical calculations and queries"""
    if not fuzzy_match(command, ["calculate", "math", "solve"]):
        return False
    
    query = command.replace("calculate", "").replace("math", "").replace("solve", "").strip()
    if not query:
        speak("What would you like me to calculate?", 'normal')
        query = listen()
    if query:
        try:
            res = wolfram_client.query(query)
            answer = next(res.results).text
            speak(f"The answer is {answer}", 'normal')
        except Exception as e:
            print(f"Math error: {e}")
            speak("Couldn't solve that math problem", 'high')
    return True

def handle_joke_commands(command):
    """Handle joke requests"""
    if not fuzzy_match(command, ["joke", "funny", "make me laugh"]):
        return False
    
    try:
        joke = pyjokes.get_joke()
        speak(joke, 'normal')
    except Exception as e:
        print(f"Joke error: {e}")
        speak("Why don't scientists trust atoms? Because they make up everything!", 'normal')
    return True

def handle_time_commands(command):
    """Handle time and date queries"""
    if not fuzzy_match(command, ["time", "date", "day"]):
        return False
    
    now = datetime.datetime.now()
    if "time" in command:
        speak(f"The current time is {now.strftime('%I:%M %p')}", 'normal')
    elif "date" in command:
        speak(f"Today's date is {now.strftime('%B %d, %Y')}", 'normal')
    elif "day" in command:
        speak(f"Today is {now.strftime('%A')}", 'normal')
    else:
        speak(f"It's {now.strftime('%A, %B %d, %Y. The time is %I:%M %p')}", 'normal')
    return True

def handle_reminder_commands(command):
    """Handle reminder creation"""
    if not fuzzy_match(command, ["reminder", "remind me"]):
        return False
    
    try:
        speak("What should I remind you about?", 'normal')
        reminder = listen()
        speak("When should I remind you? For example, say 'in 30 minutes' or 'at 5 pm'", 'normal')
        when = listen()
        
        if reminder and when:
            try:
                # Parse time string
                if "at" in when:
                    # Specific time
                    time_str = when.replace("at", "").strip()
                    reminder_time = datetime.datetime.strptime(time_str, "%I %p")
                    now = datetime.datetime.now()
                    reminder_time = reminder_time.replace(year=now.year, month=now.month, day=now.day)
                    if reminder_time < now:
                        reminder_time += datetime.timedelta(days=1)
                else:
                    # Relative time
                    nums = [int(s) for s in when.split() if s.isdigit()]
                    if not nums:
                        delta = datetime.timedelta(minutes=30)  # Default 30 minutes
                    else:
                        num = nums[0]
                        if "hour" in when:
                            delta = datetime.timedelta(hours=num)
                        elif "day" in when:
                            delta = datetime.timedelta(days=num)
                        elif "week" in when:
                            delta = datetime.timedelta(weeks=num)
                        else:
                            delta = datetime.timedelta(minutes=num)  # Default minutes
                    reminder_time = datetime.datetime.now() + delta
                
                # Calculate delay in seconds
                delay = (reminder_time - datetime.datetime.now()).total_seconds()
                
                # Create and start reminder thread
                def create_reminder():
                    time.sleep(delay)
                    notification.notify(
                        title="Reminder",
                        message=reminder,
                        timeout=10
                    )
                    speak(f"Reminder: {reminder}", 'high')
                
                Thread(target=create_reminder).start()
                speak(f"I'll remind you at {reminder_time.strftime('%I:%M %p')}", 'normal')
            except Exception as e:
                print(f"Reminder time parsing error: {e}")
                speak("I couldn't understand that time. Please try again.", 'high')
    except Exception as e:
        print(f"Reminder error: {e}")
        speak("Couldn't set that reminder", 'high')
    return True

def handle_timer_commands(command):
    """Handle timer creation"""
    if not fuzzy_match(command, ["timer", "countdown"]):
        return False
    
    try:
        speak("How long should the timer be? For example, say '5 minutes' or '1 hour'", 'normal')
        duration = listen()
        
        if duration:
            # Parse duration
            nums = [int(s) for s in duration.split() if s.isdigit()]
            if not nums:
                delta = datetime.timedelta(minutes=1)  # Default 1 minute
            else:
                num = nums[0]
                if "hour" in duration:
                    delta = datetime.timedelta(hours=num)
                elif "second" in duration:
                    delta = datetime.timedelta(seconds=num)
                else:
                    delta = datetime.timedelta(minutes=num)  # Default minutes
            
            # Start timer
            def run_timer():
                remaining = delta.total_seconds()
                while remaining > 0:
                    mins, secs = divmod(remaining, 60)
                    hours, mins = divmod(mins, 60)
                    timer_str = f"{int(hours):02d}:{int(mins):02d}:{int(secs):02d}"
                    print(f"⏳ Timer: {timer_str}", end='\r')
                    time.sleep(1)
                    remaining -= 1
                print()
                notification.notify(
                    title="Timer Complete",
                    message=f"{duration} timer has finished",
                    timeout=10
                )
                speak("Timer complete!", 'high')
            
            Thread(target=run_timer).start()
            speak(f"Timer set for {duration}", 'normal')
    except Exception as e:
        print(f"Timer error: {e}")
        speak("Couldn't set that timer", 'high')
    return True

def handle_alarm_commands(command):
    """Handle alarm creation"""
    if not fuzzy_match(command, ["alarm", "wake me"]):
        return False
    
    try:
        speak("What time should I set the alarm for? For example, say '7 am' or '11:30 pm'", 'normal')
        alarm_time = listen()
        
        if alarm_time:
            # Parse alarm time
            try:
                if ":" in alarm_time:
                    alarm_time = datetime.datetime.strptime(alarm_time, "%I:%M %p")
                else:
                    alarm_time = datetime.datetime.strptime(alarm_time, "%I %p")
                
                now = datetime.datetime.now()
                alarm_time = alarm_time.replace(year=now.year, month=now.month, day=now.day)
                if alarm_time < now:
                    alarm_time += datetime.timedelta(days=1)
                
                # Calculate delay in seconds
                delay = (alarm_time - now).total_seconds()
                
                # Create and start alarm thread
                def create_alarm():
                    time.sleep(delay)
                    for _ in range(5):  # Alarm sound loop
                        notification.notify(
                            title="Alarm",
                            message="Time to wake up!",
                            timeout=10
                        )
                        winshell.Beep(1000, 1000)  # Frequency, duration
                        time.sleep(1)
                    speak("Wake up! Your alarm is going off!", 'high')
                
                Thread(target=create_alarm).start()
                speak(f"Alarm set for {alarm_time.strftime('%I:%M %p')}", 'normal')
            except ValueError:
                speak("I couldn't understand that time format. Please try again.", 'high')
    except Exception as e:
        print(f"Alarm error: {e}")
        speak("Couldn't set that alarm", 'high')
    return True

# ===================== MAIN COMMAND PROCESSING =====================

def process_command(command):
    """Process the user command through all available handlers"""
    global assistant_active
    
    if not command:
        return
    
    # Check for exit commands
    exit_commands = ["exit", "quit", "goodbye", "stop", "sleep"]
    if any(exit_cmd in command for exit_cmd in exit_commands):
        assistant_active = False
        speak("Goodbye Giri. Let me know if you need anything else.", 'high')
        return
    
    # Check for thank you
    if "thank you" in command:
        responses = [
            "You're welcome Giri",
            "Happy to help",
            "Anytime",
            "My pleasure",
            "Glad I could assist you"
        ]
        speak(random.choice(responses), 'normal')
        return
    
    # Check for greeting
    greetings = ["hello", "hi", "hey"]
    if any(greet in command for greet in greetings):
        speak("Hello Giri, how can I help you?", 'normal')
        return
    
    # Check for how are you
    if "how are you" in command:
        speak("I'm functioning optimally, thank you for asking. How can I assist you today?", 'normal')
        return
    
    # Check for what's your name
    if "your name" in command:
        speak("I'm Jarvis, your personal assistant. At your service Giri.", 'normal')
        return
    
    # Check for what can you do
    if "what can you do" in command or "help" in command:
        speak("I can perform many tasks including web searches, system controls, media playback, "
              "setting reminders, answering questions, and much more. Just ask and I'll do my best to help.", 'normal')
        return
    
    # Process command through all handlers
    handlers = [
        handle_media_control_commands,
        handle_system_commands,
        handle_app_commands,
        handle_web_search_commands,
        handle_weather_commands,
        handle_news_commands,
        handle_wikipedia_commands,
        handle_email_commands,
        handle_math_commands,
        handle_joke_commands,
        handle_time_commands,
        handle_reminder_commands,
        handle_timer_commands,
        handle_alarm_commands,
        handle_youtube_commands,
        handle_whatsapp_commands,
        handle_linkedin_commands,
        handle_chatgpt_commands,
        handle_notepad_commands,
        handle_calculator_commands,
        handle_file_explorer_commands,
        handle_chrome_commands,
        handle_google_commands,
        handle_cmd_commands
    ]
    
    for handler in handlers:
        if handler(command):
            return
    
    # If no handler processed the command
    speak("I'm not sure how to help with that. Could you try rephrasing?", 'normal')

# ===================== MAIN LOOP =====================


def main():
    """Main execution loop for the assistant"""
    global running, assistant_active
    
    # Initial setup
    adjust_for_ambient_noise()
    speak("Initialization complete. Systems are now operational.", 'high')
    
    # Main loop
    while running:
        try:
            if assistant_active:
                command = listen()
                if command:
                    process_command(command)
            else:
                command = listen_for_wake_word()
                if command:
                    process_command(command)
        except KeyboardInterrupt:
            running = False
            speak("Shutting down all systems. Goodbye Giri.", 'high')
        except Exception as e:
            print(f"Main loop error: {e}")
            time.sleep(1)
def handle_command(cmd):
    speak(f"Handling command: {cmd}", 'normal')
    # Add logic to parse cmd and trigger features
    return f"Handled command: {cmd}"


if __name__ == "__main__":
    main()