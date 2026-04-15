// Load our library that generates the document
import Docxtemplater from "docxtemplater";
import expressionParser from "docxtemplater/expressions.js";
// Load PizZip library to load the docx/pptx/xlsx file in memory
import PizZip from "pizzip";

// Builtin file system utilities
import fs from "fs";
import path from "path";

// Load the input data
import input from "../inputs/people.json";

// Load the docx file as binary content
const content = fs.readFileSync(
  path.resolve(__dirname, "../templates/list-punctuation-version-e.docx"),
  "binary",
);

// Unzip the content of the file
const zip = new PizZip(content);

/*
 * Parse the template.
 * This function throws an error if the template is invalid,
 * for example, if the template is "Hello {user" (missing closing tag)
 */

/*
 * Punctuation styling helpers
 */
type PunctuationOption = {
  style?: string | null;
  sep?: string | null;
  conj?: string | null;
  isOxford: boolean;
};

const PUNC_STYLES: Record<string, PunctuationOption> = {
  SEMICOLON_AND: { sep: ";", conj: "and", isOxford: true },
  COMMA_AND: { sep: ",", conj: "and", isOxford: true },
};

function makePuncEvaluator(context: Docxtemplater.DXT.ParserContext) {
  const totalLength = context.scopePathLength.at(-1);
  const index = context.scopePathItem.at(-1);
  if (totalLength === undefined || index === undefined) {
    throw new SyntaxError("$punc must be used inside a loop");
  }

  return (
    style: string,
    sep: string,
    conj: string,
    isOxford: boolean = true,
  ) => {
    const isLast = index === totalLength - 1;
    if (isLast) return "";

    const isSecondToLast = index === totalLength - 2;
    const isPair = totalLength === 2;

    // If `style` parameter is given, look it up using PUNC_STYLES and
    // ignore any other parameters that may have been provided.
    if (style) {
      const preset = PUNC_STYLES[style];
      if (!preset) {
        throw new SyntaxError(
          `Unknown $punc style provided: "${style}". Valid styles are: ${Object.keys(PUNC_STYLES).join(", ")}`,
        );
      }
      ({ sep, conj, isOxford } = preset);
    }

    // Punctuation decisions
    // Two-item list: if we have a conjunction, it replaces the separator entirely
    if (isPair) {
      return conj ? ` ${conj} ` : `${sep} `;
    }

    // Three-or-more list, slot right before the last item
    if (isSecondToLast) {
      if (!conj) return `${sep} `;
      if (isOxford) return `${sep} ${conj} `;
      return ` ${conj} `;
    }

    // Any earlier slot: just the separator
    return `${sep} `;
  };
}

const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
  parser: expressionParser.configure({
    evaluateIdentifier(tag, scope, scopeList, context) {
      if (tag === "$punc") return makePuncEvaluator(context);
    },
  }),
});

doc.render(input);

// Save the rendered document to a file
const buf = doc.toBuffer();
fs.writeFileSync(
  path.resolve(__dirname, "../outputs/output-punctuation-e.docx"),
  buf,
);
