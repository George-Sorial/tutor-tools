
function createLessonSlides(topic, studentName, year, templateId) {
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(`Lesson: ${topic} for ${studentName} (${year})`);
    const presentation = SlidesApp.openById(newFile.getId());
    
    const aiContent = generateContent(topic, year);
    const sections = parseAIContent(aiContent);
    
    // Replace FIRST slide with topic/student info
    const firstSlide = presentation.getSlides()[0];
    firstSlide.replaceAllText('{{CONTENT}}', sections[0] || '');
    Object.entries({
      '{{TOPIC}}': topic,
      '{{STUDENT_NAME}}': studentName,
      '{{YEAR}}': year
    }).forEach(([placeholder, value]) => {
      firstSlide.replaceAllText(placeholder, value);
    });
    
    // Get the template slide (slide index 1) - DO NOT MODIFY
    const templateSlideIndex = 1;
    const templateSlide = presentation.getSlides()[templateSlideIndex];
    
    // Counter-based loop through remaining sections
    let counter = 1;  // Counter starts at 1 (we already did section 0)
    
    while (counter < sections.length) {
      // Duplicate the template slide
      const newSlide = presentation.appendSlide(templateSlide);
      
      // Replace {{CONTENT}} with CURRENT counter section ONLY in new slide
      newSlide.replaceAllText('{{CONTENT}}', sections[counter]);
      
      console.log(`Slide ${counter + 1}: Added section ${counter} (${sections[counter].substring(0, 30)}...)`);
      
      counter++;  
    }
    
    presentation.saveAndClose();
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return `✅ Success! Created ${sections.length} slides:\n${newFile.getUrl()}`;
  } catch (error) {
    return `❌ Error: ${error.message}`;
  }
}

function parseAIContent(aiContent) {
  let cleaned = aiContent.replace(/\[\d+\]/g, '');
  const sections = cleaned.split('---SPLIT---').map(s => s.trim()).filter(s => s);
  
  console.log(`Parsed ${sections.length} sections total`);
  sections.forEach((sec, i) => {
    console.log(`Section ${i}: ${sec.substring(0, 50)}...`);
  });
  
  return sections;
}

function generateContent(topic, year) {
  const prompt = `You are designing a 60-minute KS3 Year 8 maths lesson on TOPIC = "${topic}" for a UK student.

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
- 3 reflection prompts (e.g. "Explain how you would find 17% of a number without a calculator").

Important formatting rules:
- Use plain text bullet points only.
- No markdown headings like ##.
- No LaTeX.
- Do NOT include any answers or solutions.
- Keep language simple and age-appropriate for Year ${year}.`;

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
    
    // Debug: log how many sections we got
    const sections = content.split('---SPLIT---').filter(s => s.trim());
    console.log(`Generated ${sections.length} sections from AI`);
    
    return content;
  } catch (error) {
    console.log('API error:', error);
    return 'Error generating content';
  }
}

function testCreateSlides() {
  const result = createLessonSlides('Percentages', ' Student Name', '10', 'blank');
  Logger.log(result);
}

