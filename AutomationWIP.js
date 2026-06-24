function processAllStudentsFromSheet() {
  const sheetId = 'YOUR_SHEET_ID_HERE'; 
  const templateId = 'blank'; // Your Slide Template ID
  
  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  const data = sheet.getDataRange().getValues(); // Gets all data in the sheet
  
  // Loop starts at 1 to skip the Header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const studentName = row[0];      // Column A
    const year = row[1];             // Column B
    const subject = row[2];          // Column C (Used as the Topic)
    const predictedGrade = row[3];   // Column D
    const examBoard = row[4];        // Column E
    
    // Safety check: skip empty rows
    if (!studentName || !subject) continue; 
    
    console.log(`\n--- Starting processing for: ${studentName} ---`);
    console.log(`Details: Year ${year}, ${subject}, Target Grade: ${predictedGrade}, Board: ${examBoard}`);
    
    // Pass all 5 pieces of data into the slide creator
    const result = createLessonSlides(subject, studentName, year, predictedGrade, examBoard, templateId);
    console.log(result);
  }
}

/**
 * 2. THE UPDATED SLIDE CREATOR
 */
function createLessonSlides(topic, studentName, year, predictedGrade, examBoard, templateId) {
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(`Lesson: ${topic} for ${studentName} (Year ${year})`);
    const presentation = SlidesApp.openById(newFile.getId());
    
    // Pass the new variables to the AI
    const aiContent = generateContent(topic, year, predictedGrade, examBoard);
    const sections = parseAIContent(aiContent);
    
    // Replace FIRST slide with topic/student info
    const firstSlide = presentation.getSlides()[0];
    firstSlide.replaceAllText('{{CONTENT}}', sections[0] || '');
    
    // Added placeholders for Grade and Board so you can show them on the title slide if you want
    Object.entries({
      '{{TOPIC}}': topic,
      '{{STUDENT_NAME}}': studentName,
      '{{YEAR}}': year,
      '{{GRADE}}': predictedGrade,
      '{{BOARD}}': examBoard
    }).forEach(([placeholder, value]) => {
      firstSlide.replaceAllText(placeholder, value.toString());
    });
    
    // Get the template slide (slide index 1) - DO NOT MODIFY
    const templateSlideIndex = 1;
    const templateSlide = presentation.getSlides()[templateSlideIndex];
    
    // Counter-based loop through remaining sections
    let counter = 1;  // Counter starts at 1 so we don't repeat the first section
    
    while (counter < sections.length) {
      const newSlide = presentation.appendSlide(templateSlide);
      newSlide.replaceAllText('{{CONTENT}}', sections[counter]);
      console.log(`Slide ${counter + 1}: Added section ${counter} (${sections[counter].substring(0, 30)}...)`);
      counter++;  
    }
    
    presentation.saveAndClose();
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return `✅ Success! Created slides for ${studentName}:\n${newFile.getUrl()}`;
  } catch (error) {
    return `❌ Error creating slides for ${studentName}: ${error.message}`;
  }
}

/**
 * 3. PARSE OUTPUT
 */
function parseAIContent(aiContent) {
  let cleaned = aiContent.replace(/\[\d+\]/g, '');
  const sections = cleaned.split('---SPLIT---').map(s => s.trim()).filter(s => s);
  
  console.log(`Parsed ${sections.length} sections total`);
  return sections;
}

/**
 * 4. THE UPDATED PROMPT
 */
function generateContent(topic, year, predictedGrade, examBoard) {
  // The prompt now uses the Predicted Grade and Exam Board to tailor the lesson!
  const prompt = `You are designing a 60-minute maths lesson on TOPIC = "${topic}" for a UK student.

STUDENT PROFILE:
- Year Group: ${year}
- Target Grade: ${predictedGrade} (on the GCSE 1-9 scale)
- Exam Board: ${examBoard}
Ensure the difficulty, terminology, and challenge level are strictly tailored to a Grade ${predictedGrade} student studying under the ${examBoard} specification.

Structure the lesson into the following clearly separated sections, each ending with the delimiter ---SPLIT---:

[1] Learning Objectives & Context
- 2–3 bullet-point objectives starting with "Be able to…".
- 2 real-life contexts where this skill is useful.

---SPLIT---

[2] Conceptual Definition
- Intuitive explanation linking to fractions and decimals.
- 3 simple, concrete examples using small numbers.

---SPLIT---

[3] Teacher Modelling ("I do")
- Brief explanation of the main methods to use.
- 3 fully worked examples, step-by-step, but DO NOT include final answers here (just show working structure).

---SPLIT---

[4] Guided Practice ("We do")
- 4 practice questions with short prompts for what the student should think about.
- No answers.

---SPLIT---

[5] Independent Practice ("You do")
- 6 questions, mixed: bare number, word problems, and one "tricky" misconception question.
- No answers.

---SPLIT---

[6] Application / Challenge
- 1 or 2 richer problems or puzzles that combine ideas, suitable for discussion.
- No answers.

---SPLIT---

[7] Plenary / Reflection
- 2 quick-check questions.
- 3 reflection prompts.

Important formatting rules:
- Use plain text bullet points only.
- No markdown headings like ##.
- No LaTeX.
- Do NOT include any answers or solutions.`;

  try {
    const response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'sonar',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 2000 
      }),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    const content = data.choices[0].message.content;
    return content;
  } catch (error) {
    console.log('API error:', error);
    return 'Error generating content';
  }
}