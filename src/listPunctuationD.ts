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
  path.resolve(__dirname, "../templates/list-punctuation-version-d.docx"),
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
  isOxford?: boolean;
};

const PUNC_STYLES: Record<string, PunctuationOption> = {
  "semicolon and": { sep: ";", conj: "and", isOxford: true },
  "comma and": { sep: ",", conj: "and", isOxford: true },
};

const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
  delimiters: { start: "[[", end: "]]" },
  parser: expressionParser.configure({
    evaluateIdentifier(tag, scope, scopeList, context) {
      if (tag === "$punc") {
        // Capture context in the closure so the returned function can use it
        const totalLength = context.scopePathLength.at(-1);
        const index = context.scopePathItem.at(-1);

        return (opts: PunctuationOption = {}) => {
          // If `style` paramater is given, look it up using PUNC_STYLES and
          // ignore any other parmaters that may have been provided.
          if (opts.style) {
            const preset = PUNC_STYLES[opts.style];
            if (!preset) {
              throw new SyntaxError(
                `Unknown $punc style provided: "${opts.style}". Valid styles are: ${Object.keys(PUNC_STYLES).join(", ")}.`,
              );
            }
            opts = PUNC_STYLES[opts.style];
          }

          //   const sep = resolved.sep ?? null;
          //   const conj = resolved.conj ?? null;
          //   const isOxford = resolved.isOxford ?? true; // Defaults to true, as that is most common
          opts.isOxford = opts.isOxford ?? true; // Defaults to true if not provided, as that is the most common use case.
          console.log(opts, typeof opts);
          //   console.log(resolved, typeof resolved);
          console.log("---------------");

          const isLast = index === totalLength - 1;
          const isSecondToLast = index === totalLength - 2;
          const isPair = totalLength === 2;

          if (isLast) return "";

          // Two-item list: if we have a conjunction, it replaces the separator entirely
          if (isPair) {
            return opts.conj ? ` ${opts.conj} ` : `${opts.sep} `;
          }

          // Three-or-more list, slit right before the last item
          if (isSecondToLast) {
            if (!opts.conj) return `${opts.sep} `;
            if (opts.isOxford) return `${opts.sep} ${opts.conj} `;
            return ` ${opts.conj} `;
          }

          // Any earlier slot: just the separator
          return `${opts.sep} `;
        };
      }
    },
  }),
});

/*
 * Render the document : Replaces :
 * - {first_name} with John
 * - {last_name} with Doe,
 * ...
 */
doc.render(input);

/*
 * Get the output document and export it as a Node.js buffer
 * This method is available since docxtemplater@3.62.0
 */
const buf = doc.toBuffer();

// Write the Buffer to a file
fs.writeFileSync(
  path.resolve(__dirname, "../outputs/output-punctuation-d.docx"),
  buf,
);
