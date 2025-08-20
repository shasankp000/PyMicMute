# PyMicMute
A simple master microphone muting/unmuting app which supports custom keybinds and also has in-built notification support for mute/unmute events.

# Setup instructions

1. Download the exe file.
2. Run it.
3. Go to the system tray and you will find a green microphone icon.

<img width="166" height="93" alt="image" src="https://github.com/user-attachments/assets/e6ab65db-d7f9-411c-a6c8-e39f88d3f72c" />

4. Click on Settings.

<img width="420" height="607" alt="image" src="https://github.com/user-attachments/assets/1808eddf-c1e0-4512-b8c1-556d366178f7" />

5. From the dropdown menu, select the input device of your choice.
6. Click on Rebind Hotkey -> press whatever key combo / single button you feel like.
7. If you want to make sure the app starts on startup, click on the Run at startup option
8. Set the theme as per choices.
9. You can check mic mute/unmute settings by clicking on the Toggle Mic button or pressing the keybind

# Features

1. Simple to use
2. Cool UI (for a python app that is)
3. App has a resilient settings feature which works against input device crashes or power cuts, after such an event when the app is restarted or the input device is re-connected, the app will remember what state the mic was in previously and automatically mute or un-mute the mirophone.
4. Works across any meeting app/multiplayer game that uses a mic.


> Note: I made this app for a friend to solve a problem, but anyone can now use it :)

---
# Important 

The mic tray icon will **not sync** with any external meeting app / any in-game multiplayer microphone state. This app is a master control on the device's microphone passthrough, so only pressing the correct key combo will reflect on the system tray and the microphone will be blocked, blocking input audio to any and all apps using the specific device. For example pressing mute on google meet, zoom or discord will make this app mute it's mic, or pressing the key combo will not make these apps press their mute buttons those are governed by the respective apps. 

TLDR: **This app acts independently of any and all multiplayer games / video conference /  call apps so those apps will not affect this app's mic state, but this app will affect the mic of all the other apps using the selected device.**
