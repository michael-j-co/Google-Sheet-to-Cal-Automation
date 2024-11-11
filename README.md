# Google Sheets to Google Calendar Automation

## Overview
This project automates the process of converting a work schedule from Google Sheets into Google Calendar events using **Google Apps Script**. It automatically selects the most recent schedule sheet, checks for existing events to prevent duplicates, and updates the calendar accordingly.

## Features
- Automatically detects the latest work schedule sheet.
- Adds shifts to a dedicated Google Calendar.
- Handles shifts that cross midnight.
- Prevents duplicate events using unique identifiers.
- Configurable for any user’s name and calendar.

## Prerequisites
- Access to the work schedule Google Sheet (view-only is sufficient).
- A Google Calendar where the events will be created.
- Google Apps Script enabled in your Google account.

## Setup
1. Copy the `main.js` code to your own Google Apps Script project.
2. Replace the constants:
   - `SPREADSHEET_ID`: The ID of your Google Sheet.
   - `USER_NAME`: Your name as it appears in the schedule.
   - `CALENDAR_NAME`: The name of your Google Calendar.
3. Set up a daily time-driven trigger for the `checkForNewSheetAndUpdateCalendar` function.

## Usage
Once set up, the script will:
- Check for the most recent schedule sheet in Google Sheets.
- Add new shifts to your Google Calendar.
- Skip events that have already been created to avoid duplicates.

## Example Output
- Event created: "Monday Shift - Fall '24 Week 7" from 8:45 AM to 12:00 PM at Gayley Heights.
- Event already exists: "Friday Shift - Fall '24 Week 7".
