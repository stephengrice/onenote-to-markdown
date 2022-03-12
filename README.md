# onenote-to-markdown

## Pre-requisites

- You must have Pandoc installed on your system and accessible in your system PATH. Check by running `pandoc --version` in a PowerShell or cmd window.
- Python version only: You must have Python 3 installed on your system. This README assumes it is accessible via `python`, but other aliases, such as `python3`, are fine.

## Python Version

**Video coming soon**

1. Install the requirements.

```bash
pip install -r requirements.txt
```

2. Open OneNote and ensure the notebook you'd like to convert is open.

3. Run the script.

```bash
python convert.py
```

4. Find your converted notes in `~/Desktop/OneNoteExport`. You're free to do what you want with the notes now - commit them to Git, open them in [obsidian](https://obsidian.md), or whatever.

## PowerShell Version

**Video/blog post: https://pagekeysolutions.com/blog/misc/onenote-to-markdown**

This version is a bit more finnicky. It may fail, but feel free to try! Open PowerShell, navigate to this repo, and run:

```ps1
.\convert3.ps1
```