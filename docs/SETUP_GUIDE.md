# Setup Guide

## Step 1
Create Google Form using fields listed in FORM_FIELDS.md

## Step 2
Link the form to a Google Sheet

## Step 3
Open:
Extensions → Apps Script

Paste contents of `script.gs`

Replace the placeholders from the script to actual IDs.

## Step 4
Create two Google Docs:
- Room Invoice Template with Placeholders.
- Food Invoice Template with Placeholders

Copy content from `/templates/*.md` or create as per your need.
(Refer screenshots: Room-Invoice-Template-Sample.jpg, Food-Invoice-Template-Sample.jpg)

Replace tables using Google Docs tables

## Step 5
Copy template file IDs into script (Added in Step 3):
ROOM_TEMPLATE_FILE_ID
FOOD_TEMPLATE_FILE_ID

## Step 6
Add Trigger:
On form submit → createDocFromForm

Done ✅
