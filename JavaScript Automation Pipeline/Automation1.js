function createLessonSlides(topic, studentName, year, templateId) {
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(`Lesson: ${topic} for ${studentName} (${year})`);
    const presentation = SlidesApp.openById(newFile.getId());
    
    const aiContent = generateContent(topic, year);
    const sections = parseAIContent(aiContent);
    
    // Replace first slide
    const firstSlide = presentation.getSlides()[0];
    Object.entries({
      '{{TOPIC}}': topic,
      '{{STUDENT_NAME}}': studentName,
      '{{YEAR}}': year,
      '{{CONTENT}}': sections[0] || ''
    }).forEach(([placeholder, value]) => {
      presentation.replaceAllText(placeholder, value);
    });
  
    for (let i = 1; i < sections.length; i++) {
      const newSlide = presentation.appendSlide(presentation.getLayouts()[1]);
      const textBox = newSlide.insertTextBox(sections[i], 50, 50, 860, 450);
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
  return sections;
}

function generateContent(topic, year) {
  const prompt = `CRITICAL INSTRUCTION: Separate sections with EXACT TEXT: "---SPLIT---"

Year ${year} lesson on ${topic} UK:

## Definition  
Brief definition only.
---SPLIT---

## 3 Key Examples  
Exactly 3 examples, no solutions.
---SPLIT---

## 2 Practice Problems
Exactly 2 problems only - NO answers.
---SPLIT---

MUST use ---SPLIT--- exactly between sections. Bullet points only.`;

  try {
    const response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'sonar',  // Only sonar
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 400
      }),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    return data.choices[0].message.content;
  } catch (error) {
    return 'Error generating content';
  }
}

function testCreateSlides() {
  const result = createLessonSlides('Percentages', ' Student Name', '10', 'blank');
  Logger.log(result);
}