# Lecture Capture

Turn a video lecture into ordered notes with one keypress per slide.

Start a session, press a shortcut whenever the teacher shows something worth keeping, then stop. Lecture Capture saves your screenshots, in order, as a **PowerPoint, PDF, Word document, or ZIP of images**. Add your own handwritten or typed notes in any editor afterwards.

Built for students who study from YouTube and recorded lectures (JEE, NEET, GATE and similar) and don't have time to write everything by hand.

<!-- Add a screenshot of the app here:
![Lecture Capture screenshot](docs/screenshot.png)
-->

## Features

- **One-key capture:** a global shortcut (default `F8`) works while the video is in focus.
- **Always in order:** the app numbers every screenshot itself, so no sorting folders by hand.
- **Undo a mistake:** a second shortcut (default `F9`) deletes the last screenshot.
- **Capture only what matters:** drag a box over the lecture area to skip the browser bar, tabs and taskbar.
- **Four export formats:** PowerPoint (`.pptx`), PDF, Word (`.docx`) and a ZIP of images. Tick any combination.
- **Room for your own notes:** optionally add a blank note page after every screenshot.
- **Clean results:** temporary screenshots are deleted automatically once your files are saved.
- **Your file name, your folder:** a Save As window appears when you export.
- **Duplicate skipping (optional):** ignore a screenshot that looks the same as the previous one.
- **Custom shortcuts:** click a key button, press your combination, and save. It applies immediately, even mid-lecture.
- **Multi-monitor support:** choose which monitor to capture.
- **Remembers your settings** between sessions.

## Requirements

- Windows 10 or 11
- Python 3.8 or newer

Python libraries (installed in the steps below): `customtkinter`, `mss`, `keyboard`, `python-pptx`, `python-docx`, `pillow`.

## Installation

1. **Install Python** from [python.org/downloads](https://www.python.org/downloads/). On the first installer screen, tick **Add python.exe to PATH**.
2. **Download this project** (green *Code* button, then *Download ZIP*) and unzip it, or clone it:
   ```
   git clone https://github.com/<your-username>/lecture-capture.git
   cd lecture-capture
   ```
3. **Install the libraries** (one time only). Open a Command Prompt in the project folder (click the folder's address bar, type `cmd`, press Enter) and run:
   ```
   python -m pip install -r requirements.txt
   ```
4. **Run the app:**
   ```
   python lecture_capture.py
   ```

If Windows says `python is not recognized`, use `py` instead of `python` in steps 3 and 4.

## How to use

1. Type a **lecture name** (it becomes the suggested file name) and choose your **monitor**.
2. Optional: click **Select region** and drag a box over the lecture area. Click **Full screen** to go back.
3. Tick the **export formats** you want.
4. Press **Start Lecture**.
5. Press the **capture key** (default `F8`) whenever you want a screenshot. Press `F9` to delete the last one.
6. Press **Stop & Export**, choose a name and folder, and you're done.

### Options

| Option | What it does |
|---|---|
| Blank note page after each screenshot | Adds an empty slide/page after every screenshot for your own notes (not used for ZIP) |
| Skip near-identical screenshots | Ignores a capture that looks the same as the previous one |
| Beep after each screenshot | Plays a short sound so you know the capture worked |

### Changing the shortcut keys

Click the **Capture** or **Delete last** key button, press the combination you want (for example `Ctrl+Alt+S`), then click **Save keys**. Press `Esc` while recording to cancel.

Avoid shortcuts your browser or video player already uses, such as `Ctrl+Shift+S` (Firefox screenshots) or single letters like `K`, `F` and `M` on YouTube. `F8`, `F9` and `Ctrl+Alt+<letter>` are safe choices.

## Where files go

| What | Location |
|---|---|
| Your exported notes | The folder you choose in the Save As window |
| Settings | `Documents/LectureCaptures/settings.json` |
| Temporary screenshots | `Documents/LectureCaptures/_in_progress/` (deleted after a successful export) |

If the app crashes mid-lecture, your screenshots remain in the `_in_progress` folder so you can recover them.

## Build a standalone `.exe`

So you can run it without a console window or Python installed:

```
python -m pip install pyinstaller
pyinstaller --onefile --noconsole lecture_capture.py
```

The executable appears in the `dist` folder.

## Troubleshooting

| Problem | Fix |
|---|---|
| Shortcut keys do nothing | Close the app, open Command Prompt via right-click, **Run as administrator**, and start it again |
| Screenshots are black | Turn off hardware acceleration in your browser or video player |
| A shortcut opens another program's feature | Choose a different shortcut (see above) |
| Word export says `python-docx` is needed | Run `python -m pip install python-docx` |
| `ModuleNotFoundError` on start | Run `python -m pip install -r requirements.txt` with the same Python you use to run the app |

## Project structure

```
lecture-capture/
├── lecture_capture.py   # the whole app (one file, commented for beginners)
├── requirements.txt     # Python libraries to install
└── README.md
```

The code is a single file organised into numbered sections: imports, settings and colours, helper functions (exporters and image tools), and the app window. It is commented throughout, so it's a good starting point for changes. Colours and defaults are at the top.

## Roadmap

- OCR so slide text becomes searchable
- A preview of the last capture inside the app
- Saving the capture region between sessions
- macOS and Linux support (the `keyboard` library and the Explorer integration are currently Windows-oriented)

## Contributing

Issues and pull requests are welcome. If you report a bug, please include your Windows version, Python version (`python --version`) and any error text from the console.

## License

Add a license of your choice (for example MIT) as a `LICENSE` file, and update this section.
