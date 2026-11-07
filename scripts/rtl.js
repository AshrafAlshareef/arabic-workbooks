"use strict";

/**
 * Every workbook in this project is an Arabic, right-to-left document.
 *
 * `direction: "RTL"` is not an OOXML/docx option and is silently ignored, so
 * RTL has to be requested with the real properties:
 *   - paragraph: `bidirectional: true` -> <w:bidi/>
 *   - text run:  `rightToLeft: true`   -> <w:rtl/>
 *
 * Returning document-wide defaults (instead of setting each paragraph/run by
 * hand) means every page added later is RTL automatically, even if its builder
 * forgets to opt in.
 */
const ARABIC_LANGUAGE = "ar-SA";

function rtlDocumentDefaults(run = {}) {
  return {
    paragraph: { bidirectional: true },
    run: {
      ...run,
      // Enforce these even if a caller passes an LTR value by mistake.
      rightToLeft: true,
      language: { value: ARABIC_LANGUAGE, bidirectional: ARABIC_LANGUAGE },
    },
  };
}

module.exports = { ARABIC_LANGUAGE, rtlDocumentDefaults };
