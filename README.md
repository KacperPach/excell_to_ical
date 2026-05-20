# CalSync - Excel to Calendar Converter (Client-Side Only)

Convert your Excel spreadsheets (.xls/.xlsx) to .ics calendar files instantly - entirely in your browser!

## Features

- ✅ **Pure Client-Side** - No server required, all processing happens in your browser
- 🔒 **Privacy First** - Your data never leaves your device
- 📅 **All-Day Events** - Automatically converts dates to all-day calendar events
- 🇩🇪 **German Date Support** - Optimized for German Outlook exports
- 📊 **Preview** - See your data before converting
- 🎯 **Auto-Detection** - Automatically maps common column names

## How to Use

1. **Open `index.html`** in any modern web browser (Chrome, Firefox, Safari, Edge)
2. **Drop or select** your Excel file (.xls or .xlsx)
3. **Map columns** - The tool will auto-detect common columns like "Betreff", "Beginnt am", etc.
4. **Download** - Click "Download .ics Calendar" to get your calendar file
5. **Import** - Import the .ics file into Google Calendar, Outlook, Apple Calendar, or any other calendar app

## No Installation Required

Just open `index.html` in your browser - that's it! No server, no npm install, no dependencies.

## What Changed

This version has been simplified to be completely client-side:
- ❌ Removed Node.js backend server
- ❌ Removed Express and all npm dependencies  
- ❌ Removed subscription/sync features
- ✅ Direct download of .ics files
- ✅ 100% privacy - everything stays on your device

## Compatible With

- Google Calendar
- Microsoft Outlook
- Apple Calendar
- Thunderbird
- Yahoo Calendar
- Any iCal-compatible application

## Technical Details

- Uses SheetJS (XLSX) library for Excel parsing
- Generates RFC 5545 compliant iCalendar files
- Sanitizes all input to prevent XSS/injection attacks
- Supports German date formats (e.g., "Montag, 18. Mai 2026")
