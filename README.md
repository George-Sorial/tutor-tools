# AI-Powered Lesson Slide Generator 

This project uses Google Apps Script (GAS) and the Perplexity AI API (`sonar` model) to automatically generate highly tailored, structured educational slide decks for students. 

This repository tracks the evolution of the script across three distinct versions, moving from a basic prototype to a fully automated batch-processing tool.

---

## File Structure & Version History

### Version 1: The Prototype (`Automation1.js`)
The initial proof-of-concept script. It generates a basic 3-part lesson and hardcodes the formatting directly into Google Slides.
* **AI Prompt:** Short prompt requesting 3 sections (Definition, Examples, Practice Problems). Output capped at 400 tokens.
* **Slide Generation:** Duplicates a template presentation, replacing `{{TOPIC}}`, `{{STUDENT_NAME}}`, and `{{YEAR}}` on the title slide.
* **Formatting:** Appends new slides using standard layouts and hardcodes text boxes using absolute coordinates (`insertTextBox(..., 50, 50, 860, 450)`).
* **Parsing:** Uses the `---SPLIT---` delimiter to break the AI text into an array of slide content.

### Version 2: The Template & Pedagogy Upgrade (`Automation2.js`)
A major upgrade focused on educational quality and visual styling. 
* **AI Prompt (Massive Expansion):** Shifts to a full 60-minute KS3 lesson plan structured into 7 distinct pedagogical sections (Learning Objectives, Conceptual Definition, Teacher Modelling, Guided Practice, Independent Practice, Application/Challenge, Plenary). Output capped at 2000 tokens.
* **Slide Generation (Template Slide):** Abandons hardcoded text boxes. It now duplicates a pre-formatted "Template Slide" (Slide Index 1) and replaces a `{{CONTENT}}` tag. This allows the user to style the font, size, and layout in Google Slides beforehand.
* **Debugging:** Introduces robust `console.log()` statements to track AI section generation, parsed content arrays, and slide creation loops.

### Version 3: Google Sheets Batch Automation (`AutomationWIP.js`)
The production-ready version. It integrates Google Sheets to process multiple students at once and tailors the AI pedagogy to specific exam boards and predicted grades.
* **Data Source Integration:** Uses `SpreadsheetApp` to read rows from a Google Sheet containing: Student Name, Year, Subject, Predicted Grade, and Exam Board.
* **Batch Processing:** The `processAllStudentsFromSheet()` function loops through the sheet, skipping empty rows, and triggers the slide generator for every student automatically.
* **Dynamic AI Tailoring:** The prompt now injects the student's `Target Grade` and `Exam Board` (e.g., AQA, Edexcel), instructing the AI to adjust the difficulty, terminology, and challenge level accordingly.

---

## Setup & Usage

### Prerequisites
1.  **Google Workspace:** You need a Google account to use Google Sheets, Google Slides, and Google Apps Script.
2.  **Perplexity API Key:** Required to fetch the AI-generated content.

### Implementation Steps
1.  Open [Google Apps Script](https://script.google.com/).
2.  Create a new project and paste the contents of the desired version (Version 3 recommended).
3.  Set your API Key in the `API_KEY` constant.
4.  Create a **Google Slide Template**. 
    * Slide 1 (Title): Include tags like `{{TOPIC}}`, `{{STUDENT_NAME}}`, `{{YEAR}}`, `{{GRADE}}`, and `{{BOARD}}`.
    * Slide 2 (Content Template): Include the tag `{{CONTENT}}` styled however you want the lesson text to appear.
5.  Copy the Slide Template ID from its URL and paste it into the script.
6.  **(Version 3 Only):** Create a Google Sheet with columns for Name, Year, Subject, Grade, and Board. Add the Sheet ID to the script.
7.  Run the main function (`testCreateSlides` for V1/V2, or `processAllStudentsFromSheet` for V3). 
8.  *Note: On the first run, Google will ask for authorization to access your Drive and make external API requests.*

---

## Core Functions

* `generateContent()`: Handles the HTTP `POST` request to Perplexity's API.
* `parseAIContent()`: Cleans the AI response (removes citation brackets) and splits the text into an array based on the `---SPLIT---` keyword.
* `createLessonSlides()`: Handles the Google Drive and Google Slides API interactions (copying templates, appending slides, replacing text tags, updating sharing permissions).