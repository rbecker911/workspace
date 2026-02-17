/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { docs_v1 } from 'googleapis';
import { marked } from 'marked';

interface FormatRange {
  start: number;
  end: number;
  type: 'bold' | 'italic' | 'code' | 'link' | 'heading';
  url?: string;
  headingLevel?: number;
  isParagraph?: boolean;
}

interface ParsedMarkdown {
  plainText: string;
  formattingRequests: docs_v1.Schema$Request[];
}

/**
 * Parses markdown text and generates Google Docs API requests for formatting.
 * Uses the marked library to lex markdown tokens and extract formatting.
 */
export function parseMarkdownToDocsRequests(
  markdown: string,
  startIndex: number,
): ParsedMarkdown {
  // Split markdown into lines to handle block elements like headings manually
  // This preserves the original behavior of treating lines as separate blocks
  const lines = markdown.split('\n');

  let plainText = '';
  const formattingRanges: FormatRange[] = [];
  const lexer = new marked.Lexer();

  // Helper function to process tokens recursively
  function processTokens(tokens: any[]) {
    tokens.forEach((token) => {
      const start = plainText.length;

      if (token.type === 'text' || token.type === 'escape') {
        if (token.tokens) {
          processTokens(token.tokens);
        } else {
          // marked decodes HTML entities in text, so we get the plain text directly
          plainText += token.text;
        }
      } else if (token.type === 'strong') {
        if (token.tokens) processTokens(token.tokens);
        else plainText += token.text;

        const end = plainText.length;
        formattingRanges.push({ start, end, type: 'bold' });
      } else if (token.type === 'em') {
        if (token.tokens) processTokens(token.tokens);
        else plainText += token.text;

        const end = plainText.length;
        formattingRanges.push({ start, end, type: 'italic' });
      } else if (token.type === 'codespan') {
        // codespan text is usually the raw code content
        plainText += token.text;
        const end = plainText.length;
        formattingRanges.push({ start, end, type: 'code' });
      } else if (token.type === 'link') {
        if (token.tokens) processTokens(token.tokens);
        else plainText += token.text;

        const end = plainText.length;
        formattingRanges.push({ start, end, type: 'link', url: token.href });
      } else {
        // Fallback for other tokens
        if (token.tokens) {
          processTokens(token.tokens);
        } else if (token.text) {
          plainText += token.text;
        }
      }
    });
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Check if this is a heading line
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);

    if (headingMatch) {
      const level = headingMatch[1].length;
      const content = headingMatch[2];

      const start = plainText.length;

      // Parse inline content within the heading
      const tokens = lexer.inlineTokens(content);
      processTokens(tokens);

      const end = plainText.length;

      // Mark the entire heading range
      formattingRanges.push({
        start,
        end,
        type: 'heading',
        headingLevel: level,
        isParagraph: true,
      });

    } else if (line.trim()) {
      // For non-heading, non-empty lines, use inlineTokens
      const tokens = lexer.inlineTokens(line);
      processTokens(tokens);
    } else {
      // Empty line - content is empty, just plainText remains unchanged (except for newline added below)
    }

    // Add newline after line content if not the last line
    // This matches the behavior of joining lines with '\n'
    if (i < lines.length - 1) {
      plainText += '\n';
    }
  }

  // Generate formatting requests
  const formattingRequests: docs_v1.Schema$Request[] = [];

  for (const range of formattingRanges) {
    const textStyle: docs_v1.Schema$TextStyle = {};
    const fields: string[] = [];

    if (range.type === 'bold') {
      textStyle.bold = true;
      fields.push('bold');
    } else if (range.type === 'italic') {
      textStyle.italic = true;
      fields.push('italic');
    } else if (range.type === 'code') {
      textStyle.weightedFontFamily = {
        fontFamily: 'Courier New',
        weight: 400,
      };
      textStyle.backgroundColor = {
        color: {
          rgbColor: {
            red: 0.95,
            green: 0.95,
            blue: 0.95,
          },
        },
      };
      fields.push('weightedFontFamily', 'backgroundColor');
    } else if (range.type === 'link' && range.url) {
      textStyle.link = {
        url: range.url,
      };
      textStyle.foregroundColor = {
        color: {
          rgbColor: {
            red: 0.06,
            green: 0.33,
            blue: 0.8,
          },
        },
      };
      textStyle.underline = true;
      fields.push('link', 'foregroundColor', 'underline');
    } else if (
      range.type === 'heading' &&
      range.headingLevel &&
      range.isParagraph
    ) {
      // Use updateParagraphStyle for headings as per Google Docs API best practices
      const headingStyles: { [key: number]: string } = {
        1: 'HEADING_1',
        2: 'HEADING_2',
        3: 'HEADING_3',
        4: 'HEADING_4',
        5: 'HEADING_5',
        6: 'HEADING_6',
      };

      const namedStyleType = headingStyles[range.headingLevel] || 'HEADING_1';

      // Create a separate updateParagraphStyle request for headings
      formattingRequests.push({
        updateParagraphStyle: {
          paragraphStyle: {
            namedStyleType: namedStyleType,
          },
          range: {
            startIndex: startIndex + range.start,
            endIndex: startIndex + range.end,
          },
          fields: 'namedStyleType',
        },
      });

      // Skip the normal text style formatting for headings
      continue;
    }

    if (fields.length > 0) {
      formattingRequests.push({
        updateTextStyle: {
          range: {
            startIndex: startIndex + range.start,
            endIndex: startIndex + range.end,
          },
          textStyle: textStyle,
          fields: fields.join(','),
        },
      });
    }
  }

  return {
    plainText,
    formattingRequests,
  };
}

/**
 * Handles line breaks and paragraphs in markdown text
 */
export function processMarkdownLineBreaks(text: string): string {
  // Convert double line breaks to paragraph breaks
  // Single line breaks remain as-is
  return text.replace(/\n\n+/g, '\n\n');
}
