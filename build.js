#!/usr/bin/env node
const fs = require('fs');

// Read template
let template = fs.readFileSync('presentation.md', 'utf8');

// Extract global variables from front-matter
const frontMatterMatch = template.match(/^---\n([\s\S]*?)\n---/);
const globalVars = {};

if (frontMatterMatch) {
  const frontMatter = frontMatterMatch[1];
  const varMatches = frontMatter.matchAll(/^([A-Z_]+):\s*"([^"]*)"/gm);
  for (const match of varMatches) {
    globalVars[match[1]] = match[2];
  }
}

// Process the template slide by slide
let output = template;
const slideRegex = /(<!--[\s\S]*?-->)/g;
let match;
let lastIndex = 0;
let result = '';

// Split by comment blocks and process each
const parts = template.split(/(<!--[\s\S]*?-->)/);

for (let i = 0; i < parts.length; i++) {
  let part = parts[i];
  
  if (part.startsWith('<!--')) {
    // Extract local variables from this comment
    const localVars = {...globalVars};
    const localMatches = part.matchAll(/^([A-Z_]+):\s*"([^"]*)"/gm);
    for (const match of localMatches) {
      localVars[match[1]] = match[2];
    }
    
    // Process next part (slide content) with these variables
    result += part;
    if (i + 1 < parts.length) {
      i++;
      let slideContent = parts[i];
      
      // Replace variables in this slide
      for (const [key, value] of Object.entries(localVars)) {
        const regex = new RegExp(`{{${key}}}`, 'g');
        slideContent = slideContent.replace(regex, value);
      }
      result += slideContent;
    }
  } else {
    // Not in a comment block, just replace global variables
    for (const [key, value] of Object.entries(globalVars)) {
      const regex = new RegExp(`{{${key}}}`, 'g');
      part = part.replace(regex, value);
    }
    result += part;
  }
}

// Write output
fs.writeFileSync('presentation.output.md', result);
console.log('✓ Generated presentation.output.md from source');
console.log('  Global variables:', Object.keys(globalVars).join(', '));
