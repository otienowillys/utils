const mammoth = require('mammoth');
const fs = require('fs');

async function extractHeadingsFromDocx(filePath) {
  try {
    const { value } = await mammoth.convertToHtml({ path: filePath });
    const htmlContent = value;

    // Match all heading tags from h1 to h6
    const headingRegex = /<(h[1-6])[^>]*>(.*?)<\/\1>/gi;
    const matches = [...htmlContent.matchAll(headingRegex)];

    const root = {}; // Final JSON
    const stack = []; // Track nesting

    matches.forEach((match, index) => {
      const level = parseInt(match[1].substring(1)); // Number: 1-6
      const title = match[2].replace(/<\/?[^>]+(>|$)/g, "").trim();

      const content = getContentBetweenHeadings(matches, index, htmlContent);

      const newNode = {
        content: content,
        subsections: {}
      };

      while (stack.length > 0 && stack[stack.length - 1].level >= level) {
        stack.pop(); // Move back to the correct parent
      }

      if (stack.length === 0) {
        root[title] = newNode;
      } else {
        const parent = stack[stack.length - 1].node;
        parent.subsections[title] = newNode;
      }

      stack.push({ level: level, node: newNode });
    });

    return root;
  } catch (error) {
    console.error('Error:', error);
  }
}

function getContentBetweenHeadings(matches, currentIndex, htmlContent) {
  const currentMatch = matches[currentIndex];
  const nextMatch = matches[currentIndex + 1];

  const startIdx = currentMatch.index + currentMatch[0].length;
  const endIdx = nextMatch ? nextMatch.index : htmlContent.length;

  let content = htmlContent.substring(startIdx, endIdx).trim();
  content = content.replace(/<\/?[^>]+(>|$)/g, "").replace(/\s+/g, ' ').trim(); // Clean HTML tags and whitespace
  return content;
}

// --- Main ---
(async () => {
  const filePath = 'input.docx'; // Path to your DOCX file
  const nestedJson = await extractHeadingsFromDocx(filePath);

  fs.writeFileSync('output.json', JSON.stringify(nestedJson, null, 2), 'utf-8');
  console.log('✅ JSON extracted successfully and saved to output.json');
})();












{
  "Section 1": {
    "content": "Content under Section 1",
    "subsections": {
      "Subsection 1.1": {
        "content": "Content under Subsection 1.1",
        "subsections": {
          "Detail 1.1.1": {
            "content": "Content under Detail 1.1.1",
            "subsections": {
              "Point 1.1.1.1": {
                "content": "Content under Point 1.1.1.1",
                "subsections": {
                  "Subpoint 1.1.1.1.1": {
                    "content": "Content under Subpoint 1.1.1.1.1",
                    "subsections": {
                      "Minor 1.1.1.1.1.1": {
                        "content": "Content under Minor 1.1.1.1.1.1"
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
}




Assess this document content: [DocumentA_Text]
Based on these requirements: [DocumentB_Text]

For each requirement:
- Confirm if it is met (Yes/No).
- Provide supporting text.
- Note any gaps.
- Score 0 (missing) or 1 (satisfied).
- Create a table with columns: Requirement | Found | Evidence | Gap/Issue | Score (0–1).
Calculate the overall compliance percentage.

